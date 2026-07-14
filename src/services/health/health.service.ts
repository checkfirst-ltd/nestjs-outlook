import { Injectable, Logger } from '@nestjs/common';
import { EventEmitter2 } from '@nestjs/event-emitter';
import { TenantUserService } from '../tenant/tenant-user.service';
import { MicrosoftSubscriptionService } from '../subscription/microsoft-subscription.service';
import { OutlookWebhookSubscriptionRepository } from '../../repositories/outlook-webhook-subscription.repository';
import { MicrosoftUser } from '../../entities/microsoft-user.entity';
import { OutlookWebhookSubscription } from '../../entities/outlook-webhook-subscription.entity';
import { MicrosoftUserStatus } from '../../enums/microsoft-user-status.enum';
import { MicrosoftTenantStatus } from '../../enums/microsoft-tenant-status.enum';
import { UserHealthStatus } from '../../enums/user-health-status.enum';
import { OutlookEventTypes } from '../../enums/event-types.enum';
import { mapWithConcurrency } from '../../utils/concurrent-map.util';

/** How a user authenticates — determines which subscription to check / recreate. */
export type UserAuthMode = 'delegated' | 'app-only' | 'none';

/** Options for a health check. */
export interface HealthCheckOptions {
  /** Also verify the subscription still exists at Microsoft Graph (extra Graph calls). */
  verifyAtGraph?: boolean;
}

/** Per-user health verdict. Only `HEALTHY` (and `connected: true`) means connected. */
export interface UserHealth {
  externalUserId: string;
  connected: boolean;
  status: UserHealthStatus;
  authMode: UserAuthMode;
  /** Whether the state can be auto-fixed by recreating the subscription. */
  recoverable: boolean;
  subscriptionId?: string;
  microsoftUserId?: string;
  tenantId?: string;
  reason?: string;
}

/** A health verdict plus the recovery action that was taken. */
export interface UserHealthReport extends UserHealth {
  action: 'none' | 'recreated' | 'recovery_failed' | 'reported';
  newSubscriptionId?: string;
  recoveryError?: string;
}

/** Summary of a bulk recover run. `total = healthy + recovered + unrecoverable + failed`. */
export interface HealthCheckResult {
  total: number;
  healthy: number;
  recovered: number;
  unrecoverable: number;
  failed: number;
  results: UserHealthReport[];
}

/**
 * Diagnoses and recovers user connection health across delegated and app-only auth.
 *
 * "Healthy" combines the `microsoft_users` row (exists, active, status) and its active Outlook
 * **calendar** subscription (present, not expired, receiving notifications), and — when
 * `verifyAtGraph` is set — a live Microsoft Graph check. Anything else gets a specific verdict.
 *
 * Recovery is auth-mode-aware and composes existing primitives: it recreates fixable states
 * (missing / expired / stale / gone-at-Graph) via the delegated or app-only create (both of which
 * remove a stale subscription first), and **reports** the states a human must resolve
 * (`NEEDS_REAUTH`, `NEEDS_ADMIN`, `INACTIVE`, …) instead of looping on them.
 *
 * This is a facade over the existing status model + recovery primitives — it does not add new
 * state or schedulers; the 6-hour / 3am crons keep running independently.
 */
@Injectable()
export class HealthService {
  private readonly logger = new Logger(HealthService.name);

  /** A subscription with no notification since this many hours ago is considered stale. */
  private readonly STALE_HOURS = 24;

  /** Max concurrent Graph verifications / recoveries — gentle on Graph rate limits. */
  private readonly HEALTH_CONCURRENCY = 5;

  constructor(
    private readonly tenantUserService: TenantUserService,
    private readonly subscriptionService: MicrosoftSubscriptionService,
    private readonly webhookSubscriptionRepository: OutlookWebhookSubscriptionRepository,
    private readonly eventEmitter: EventEmitter2,
  ) {}

  /**
   * Diagnose a single user's connection health.
   */
  async checkUser(externalUserId: string, opts?: HealthCheckOptions): Promise<UserHealth> {
    const [health] = await this.checkUsers([externalUserId], opts);
    return health;
  }

  /**
   * Diagnose a batch of users. Two bulk DB reads (user rows + active subscriptions); when
   * `verifyAtGraph` is set, the otherwise-healthy users are verified at Graph at bounded
   * concurrency. Results preserve input order.
   */
  async checkUsers(externalUserIds: string[], opts?: HealthCheckOptions): Promise<UserHealth[]> {
    if (externalUserIds.length === 0) {
      return [];
    }

    const rows = await this.tenantUserService.findUsersByExternalIds(externalUserIds);
    const userByExternalId = new Map(rows.map((row) => [row.externalUserId, row]));

    const userIds = rows.map((row) => row.id);
    const activeSubs = userIds.length
      ? await this.webhookSubscriptionRepository.findActiveByUserIds(userIds)
      : [];
    const calendarSubByUserId = this.indexCalendarSubscriptions(activeSubs);

    const healths = externalUserIds.map((externalUserId) => {
      const user = userByExternalId.get(externalUserId);
      const sub = user ? calendarSubByUserId.get(user.id) ?? null : null;
      return this.diagnose(externalUserId, user, sub);
    });

    if (!opts?.verifyAtGraph) {
      return healths;
    }

    // Verify the DB-healthy ones against Graph; a 404 downgrades to MISSING_AT_GRAPH.
    return mapWithConcurrency(healths, this.HEALTH_CONCURRENCY, async (health) => {
      if (health.status !== UserHealthStatus.HEALTHY || !health.subscriptionId) {
        return health;
      }
      const user = userByExternalId.get(health.externalUserId);
      const sub = user ? calendarSubByUserId.get(user.id) : undefined;
      if (!sub) {
        return health;
      }
      const graph = await this.subscriptionService.verifySubscriptionAtGraph(sub);
      if (graph === 'missing') {
        return {
          ...health,
          connected: false,
          status: UserHealthStatus.MISSING_AT_GRAPH,
          recoverable: true,
          reason: 'Subscription not found at Microsoft Graph',
        };
      }
      // 'present' or 'unknown' — keep the DB verdict (don't downgrade on an inconclusive check).
      return health;
    });
  }

  /**
   * Diagnose then recover a batch of users: auto-fix the recoverable ones (recreate the
   * subscription via their auth mode) and report the rest. Emits
   * {@link OutlookEventTypes.USER_HEALTH_RECOVERY_COMPLETED} with the summary.
   */
  async recoverUsers(externalUserIds: string[], opts?: HealthCheckOptions): Promise<HealthCheckResult> {
    const correlationId = `health-recover-${Date.now()}`;
    const healths = await this.checkUsers(externalUserIds, opts);

    const reports = await mapWithConcurrency(
      healths,
      this.HEALTH_CONCURRENCY,
      (health) => this.recoverOne(health, correlationId),
    );

    const summary: HealthCheckResult = {
      total: reports.length,
      healthy: reports.filter((r) => r.action === 'none').length,
      recovered: reports.filter((r) => r.action === 'recreated').length,
      unrecoverable: reports.filter((r) => r.action === 'reported').length,
      failed: reports.filter((r) => r.action === 'recovery_failed').length,
      results: reports,
    };

    this.logger.log(
      `[${correlationId}] Health recovery complete: ${summary.healthy} healthy, ` +
      `${summary.recovered} recovered, ${summary.unrecoverable} need attention, ${summary.failed} failed`,
    );
    this.eventEmitter.emit(OutlookEventTypes.USER_HEALTH_RECOVERY_COMPLETED, summary);

    return summary;
  }

  /** Diagnose then recover a single user. */
  async recoverUser(externalUserId: string, opts?: HealthCheckOptions): Promise<UserHealthReport> {
    const health = await this.checkUser(externalUserId, opts);
    return this.recoverOne(health, `health-recover-${Date.now()}`);
  }

  // ── internals ──────────────────────────────────────────────────────────

  /**
   * Pure DB-based diagnosis for one user. `verifyAtGraph` is applied by the caller afterwards.
   */
  private diagnose(
    externalUserId: string,
    user: MicrosoftUser | undefined,
    sub: OutlookWebhookSubscription | null,
  ): UserHealth {
    const base = { externalUserId, connected: false, authMode: 'none' as UserAuthMode, recoverable: false };

    if (!user) {
      return { ...base, status: UserHealthStatus.UNKNOWN, reason: 'No Microsoft user record' };
    }
    if (!user.isActive) {
      return { ...base, status: UserHealthStatus.INACTIVE, reason: 'User is soft-deleted (isActive=false)' };
    }

    const authMode = this.authModeOf(user);
    const microsoftUserId = user.microsoftUserId ?? undefined;
    const tenantId = user.tenant?.tenantId;
    const withIds = { ...base, authMode, microsoftUserId, tenantId };

    // Dead delegated token — needs the user to re-authenticate; not auto-recoverable.
    if (user.status === MicrosoftUserStatus.CORRUPTED) {
      return { ...withIds, status: UserHealthStatus.NEEDS_REAUTH, reason: 'Delegated token is invalid (CORRUPTED)' };
    }

    if (authMode === 'app-only') {
      const tenant = user.tenant;
      if (!tenant || !microsoftUserId) {
        return { ...withIds, status: UserHealthStatus.NOT_MAPPED, reason: 'No tenant mapping' };
      }
      if (!tenant.isActive || tenant.status !== MicrosoftTenantStatus.ACTIVE) {
        return {
          ...withIds,
          status: UserHealthStatus.NEEDS_ADMIN,
          reason: `Tenant not usable (status=${tenant.status}, isActive=${tenant.isActive})`,
        };
      }
    } else if (authMode === 'none') {
      // A row with neither a delegated token nor a tenant mapping can't own or recreate a sub.
      return { ...withIds, status: UserHealthStatus.NOT_MAPPED, reason: 'No delegated token and no tenant mapping' };
    }

    // Subscription checks.
    if (!sub) {
      return { ...withIds, status: UserHealthStatus.NO_SUBSCRIPTION, recoverable: true, reason: 'No active subscription' };
    }
    if (new Date(sub.expirationDateTime).getTime() <= Date.now()) {
      return {
        ...withIds,
        status: UserHealthStatus.SUBSCRIPTION_EXPIRED,
        recoverable: true,
        subscriptionId: sub.subscriptionId,
        reason: 'Subscription expired',
      };
    }
    if (this.isStale(sub)) {
      return {
        ...withIds,
        status: UserHealthStatus.SUBSCRIPTION_STALE,
        recoverable: true,
        subscriptionId: sub.subscriptionId,
        reason: `No notification in over ${this.STALE_HOURS}h`,
      };
    }

    return { ...withIds, connected: true, status: UserHealthStatus.HEALTHY, subscriptionId: sub.subscriptionId };
  }

  /** Recreate the subscription for a recoverable user; report the rest. Never throws. */
  private async recoverOne(health: UserHealth, correlationId: string): Promise<UserHealthReport> {
    if (health.status === UserHealthStatus.HEALTHY) {
      return { ...health, action: 'none' };
    }
    if (!health.recoverable) {
      return { ...health, action: 'reported' };
    }

    try {
      let newSubscriptionId: string | undefined;
      if (health.authMode === 'app-only' && health.tenantId && health.microsoftUserId) {
        const created = await this.subscriptionService.createAppOnlyWebhookSubscription({
          tenantId: health.tenantId,
          microsoftUserId: health.microsoftUserId,
          externalUserId: health.externalUserId,
        });
        newSubscriptionId = created.id ?? undefined;
      } else {
        const created = await this.subscriptionService.createWebhookSubscription(health.externalUserId);
        newSubscriptionId = created.id ?? undefined;
      }

      return {
        ...health,
        connected: true,
        status: UserHealthStatus.HEALTHY,
        action: 'recreated',
        newSubscriptionId,
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      this.logger.warn(`[${correlationId}] Recovery failed for ${health.externalUserId}: ${message}`);
      return { ...health, action: 'recovery_failed', recoveryError: message };
    }
  }

  /** app-only when mapped into a tenant; delegated when it holds a refresh token; else none. */
  private authModeOf(user: MicrosoftUser): UserAuthMode {
    if (user.microsoftUserId && user.tenant) {
      return 'app-only';
    }
    if (user.refreshToken) {
      return 'delegated';
    }
    return 'none';
  }

  /**
   * Stale = no notification since {@link STALE_HOURS} ago. A subscription that has never received
   * a notification uses its creation time as the reference, so freshly created subs aren't flagged.
   */
  private isStale(sub: OutlookWebhookSubscription): boolean {
    const reference = sub.lastNotificationAt ?? sub.createdAt;
    const threshold = Date.now() - this.STALE_HOURS * 60 * 60 * 1000;
    return new Date(reference).getTime() < threshold;
  }

  /** Map userId → its active calendar subscription (resource ends in `/events`), if any. */
  private indexCalendarSubscriptions(
    subs: OutlookWebhookSubscription[],
  ): Map<number, OutlookWebhookSubscription> {
    const byUser = new Map<number, OutlookWebhookSubscription>();
    for (const sub of subs) {
      if (sub.resource.endsWith('/events') && !byUser.has(sub.userId)) {
        byUser.set(sub.userId, sub);
      }
    }
    return byUser;
  }
}
