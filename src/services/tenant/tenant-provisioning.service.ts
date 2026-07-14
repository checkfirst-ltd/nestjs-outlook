import { Injectable, Logger } from '@nestjs/common';
import { EventEmitter2 } from '@nestjs/event-emitter';
import { TenantUserService } from './tenant-user.service';
import { MicrosoftSubscriptionService } from '../subscription/microsoft-subscription.service';
import { MicrosoftTenantRepository } from '../../repositories/microsoft-tenant.repository';
import { mapWithConcurrency } from '../../utils/concurrent-map.util';
import { OutlookEventTypes } from '../../enums/event-types.enum';

/**
 * One user to connect into a tenant. `email` is required because the module resolves the
 * Microsoft account by email/UPN (Graph `/users?$filter=mail eq …`) — an external id alone
 * cannot be mapped to a Microsoft user.
 */
export interface BulkConnectUserInput {
  externalUserId: string;
  /** Email or user principal name (UPN) that exists in the tenant. */
  email: string;
}

/** Per-user outcome of a bulk connect. */
export interface BulkConnectUserResult {
  externalUserId: string;
  success: boolean;
  /** True when the user was already connected to this tenant and left untouched. */
  skipped?: boolean;
  microsoftUserId?: string;
  subscriptionId?: string;
  error?: string;
}

/** Summary of a bulk connect run. `total = connected + skipped + failed`. */
export interface BulkConnectResult {
  tenantId: string;
  total: number;
  /** Newly connected (mapping upserted + subscription created). */
  connected: number;
  /** Already connected to this tenant, so left as-is (not re-created). */
  skipped: number;
  failed: number;
  results: BulkConnectUserResult[];
}

/**
 * Orchestrates connecting many users into a tenant (app-only) in one call.
 *
 * For each user it upserts the `microsoft_users` mapping and creates an app-only Outlook
 * calendar webhook subscription. Users who already have a delegated subscription are handled
 * transparently: `createAppOnlyWebhookSubscription` removes any existing calendar subscription
 * (at Microsoft and locally) before creating the new one, so a mailbox never ends up with two
 * live subscriptions.
 *
 * Designed for scale: work runs at bounded concurrency (Graph subscription creation cannot be
 * batched — Graph validates each notificationUrl at creation — so concurrency, not $batch, is
 * the lever). A per-user failure is recorded and never aborts the run.
 */
@Injectable()
export class TenantProvisioningService {
  private readonly logger = new Logger(TenantProvisioningService.name);

  /**
   * Max concurrent per-user connect flows. Kept small so we don't burst Graph (each user is a
   * lookup + optional delete + create) and so simultaneous subscription-validation callbacks
   * stay manageable. Matches the revocation concurrency used elsewhere.
   */
  private readonly CONNECT_CONCURRENCY = 5;

  constructor(
    private readonly tenantUserService: TenantUserService,
    private readonly subscriptionService: MicrosoftSubscriptionService,
    private readonly tenantRepository: MicrosoftTenantRepository,
    private readonly eventEmitter: EventEmitter2,
  ) {}

  /**
   * Connect a batch of users into a tenant: upsert each mapping and create an app-only
   * calendar subscription. Users already connected to this tenant (they have an active
   * app-only subscription) are **skipped**, so a re-run never tears down a working
   * connection. Emits {@link OutlookEventTypes.TENANT_USERS_BULK_CONNECT_COMPLETED}
   * with the summary on completion (or `..._FAILED` if the run can't start), so callers that
   * kicked this off in the background can observe the outcome.
   *
   * @param tenantId - Azure AD tenant GUID the users belong to (must be connected/active).
   * @param users - Users to connect (`{ externalUserId, email }`).
   * @returns Per-user results plus connected/failed tallies.
   */
  async connectUsers(
    tenantId: string,
    users: BulkConnectUserInput[],
  ): Promise<BulkConnectResult> {
    const correlationId = `bulk-connect-${tenantId}-${Date.now()}`;

    // Dedupe the input by externalUserId (first occurrence wins). Two entries for the same
    // user would run CONCURRENTLY through the per-user worker: both dedupe guards can observe
    // "no existing subscription" before either create lands, double-subscribing the mailbox.
    const seenIds = new Set<string>();
    users = users.filter((user) => {
      if (seenIds.has(user.externalUserId)) return false;
      seenIds.add(user.externalUserId);
      return true;
    });

    this.logger.log(
      `[${correlationId}] Bulk connect requested for ${users.length} user(s) into tenant ${tenantId}`,
    );

    // Pre-flight: resolve the tenant once. If it isn't there, fail the whole run fast with a
    // clear reason instead of failing every user with the same "tenant not found" error.
    const tenant = await this.tenantRepository.findByTenantId(tenantId);
    if (!tenant) {
      const message = `Tenant not found or not connected: ${tenantId}`;
      this.logger.error(`[${correlationId}] ${message}`);
      this.eventEmitter.emit(OutlookEventTypes.TENANT_USERS_BULK_CONNECT_FAILED, {
        tenantId,
        total: users.length,
        error: message,
      });
      throw new Error(message);
    }

    // Skip users already connected to this tenant so a re-run doesn't tear down and rebuild a
    // working connection. "Connected" = has an active app-only subscription for this tenant.
    // Two bulk queries (subscriptions by tenant + user rows by external id) — no per-user query.
    const activeSubs = await this.subscriptionService.getAppOnlySubscriptionsForTenant(tenantId);
    const connectedUserIds = new Set(activeSubs.map((sub) => sub.userId));
    const existingRows = await this.tenantUserService.findUsersByExternalIds(
      users.map((user) => user.externalUserId),
    );
    const rowByExternalId = new Map(existingRows.map((row) => [row.externalUserId, row]));

    const toProcess: BulkConnectUserInput[] = [];
    const skippedResults: BulkConnectUserResult[] = [];
    for (const user of users) {
      const row = rowByExternalId.get(user.externalUserId);
      if (row && connectedUserIds.has(row.id)) {
        skippedResults.push({
          externalUserId: user.externalUserId,
          success: true,
          skipped: true,
          microsoftUserId: row.microsoftUserId ?? undefined,
        });
      } else {
        toProcess.push(user);
      }
    }

    if (skippedResults.length > 0) {
      this.logger.log(
        `[${correlationId}] ${skippedResults.length} user(s) already connected — skipping; ` +
        `processing ${toProcess.length}`,
      );
    }

    const processed = await mapWithConcurrency(
      toProcess,
      this.CONNECT_CONCURRENCY,
      (user) => this.connectOneUser(tenantId, user, correlationId),
    );

    const results = [...skippedResults, ...processed];
    const connected = processed.filter((r) => r.success).length;
    const summary: BulkConnectResult = {
      tenantId,
      total: users.length,
      connected,
      skipped: skippedResults.length,
      failed: processed.length - connected,
      results,
    };

    this.logger.log(
      `[${correlationId}] Bulk connect complete for tenant ${tenantId}: ` +
      `${summary.connected} connected, ${summary.skipped} skipped, ${summary.failed} failed`,
    );
    this.eventEmitter.emit(OutlookEventTypes.TENANT_USERS_BULK_CONNECT_COMPLETED, summary);

    return summary;
  }

  /**
   * Connect a single user: upsert the mapping, then create the app-only subscription
   * (which first removes any existing calendar subscription for the user). Never throws —
   * a failure is captured in the returned result so the batch continues.
   */
  private async connectOneUser(
    tenantId: string,
    user: BulkConnectUserInput,
    correlationId: string,
  ): Promise<BulkConnectUserResult> {
    try {
      const mapped = await this.tenantUserService.registerUserMapping(
        tenantId,
        user.externalUserId,
        user.email,
      );

      if (!mapped.microsoftUserId) {
        throw new Error('User mapping did not resolve a Microsoft user id');
      }

      const subscription = await this.subscriptionService.createAppOnlyWebhookSubscription({
        tenantId,
        microsoftUserId: mapped.microsoftUserId,
        externalUserId: user.externalUserId,
      });

      return {
        externalUserId: user.externalUserId,
        success: true,
        microsoftUserId: mapped.microsoftUserId,
        subscriptionId: subscription.id ?? undefined,
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      this.logger.warn(
        `[${correlationId}] Failed to connect user ${user.externalUserId}: ${message}`,
      );
      return {
        externalUserId: user.externalUserId,
        success: false,
        error: message,
      };
    }
  }
}
