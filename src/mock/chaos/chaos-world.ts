import { EventEmitter2 } from '@nestjs/event-emitter';
import { MicrosoftSubscriptionService } from '../../services/subscription/microsoft-subscription.service';
import { TenantUserService } from '../../services/tenant/tenant-user.service';
import { TenantProvisioningService } from '../../services/tenant/tenant-provisioning.service';
import { HealthService } from '../../services/health/health.service';
import { TenantAuthController } from '../../controllers/tenant-auth.controller';
import { HealthController } from '../../controllers/health.controller';
import { MicrosoftAuthService } from '../../services/auth/microsoft-auth.service';
import { AppOnlyAuthService } from '../../services/auth/app-only-auth.service';
import { UserIdConverterService } from '../../services/shared/user-id-converter.service';
import { GraphRateLimiterService } from '../../services/shared/graph-rate-limiter.service';
import { OutlookWebhookSubscriptionRepository } from '../../repositories/outlook-webhook-subscription.repository';
import { MicrosoftTenantRepository } from '../../repositories/microsoft-tenant.repository';
import { MicrosoftUser } from '../../entities/microsoft-user.entity';
import { MicrosoftTenant } from '../../entities/microsoft-tenant.entity';
import { MicrosoftUserStatus } from '../../enums/microsoft-user-status.enum';
import { MicrosoftTenantStatus } from '../../enums/microsoft-tenant-status.enum';
import { MicrosoftOutlookConfig } from '../../interfaces/config/outlook-config.interface';
import { Repository } from 'typeorm';
import { ChaosEngine, ChaosRates } from './chaos-engine';
import { ChaosGraph, AxiosMockLike } from './chaos-graph';
import { ChaosDb } from './chaos-db';
import { ChaosMetrics } from './chaos-metrics';

/** One user to seed into the world (DB row, Graph directory entry, optional subscription). */
export interface SeedUser {
  externalUserId: string;
  email: string;
  /**
   * - `app-only`: mapped into the tenant (microsoftUserId + tenant FK, no delegated tokens)
   * - `delegated`: delegated tokens only (no tenant mapping)
   * - `dual`: both capabilities on the row
   * - `bare`: row exists but has neither (NOT_MAPPED)
   */
  kind: 'app-only' | 'delegated' | 'dual' | 'bare';
  /** false → no microsoft_users row at all (health verdict UNKNOWN). Default true. */
  inDb?: boolean;
  /** false → user missing from the Graph directory (lookup misses). Default true. */
  inGraphDirectory?: boolean;
  status?: MicrosoftUserStatus;
  isActive?: boolean;
  tenantKey?: string;
  /** Seed an existing calendar subscription for the user. */
  sub?: {
    mode: 'app-only' | 'delegated';
    /** DB expirationDateTime in the past. */
    expired?: boolean;
    /** lastNotificationAt long ago (stale). */
    stale?: boolean;
    /** false → subscription exists in DB but NOT at Graph (drifted). Default true. */
    presentAtGraph?: boolean;
  } | null;
}

export interface SeedTenant {
  key: string;
  tenantId: string;
  status?: MicrosoftTenantStatus;
  isActive?: boolean;
}

export interface WorldOptions {
  seed: number;
  tenants?: SeedTenant[];
  users?: SeedUser[];
  graphRates?: ChaosRates;
  graphLatencyMs?: { min: number; max: number };
  dbLatencyMs?: { min: number; max: number };
}

export interface CapturedEvent {
  event: string;
  payload: unknown;
}

const HOUR_MS = 3600 * 1000;

/**
 * Builds a complete chaos world: REAL services (provisioning, subscription, tenant-user,
 * health) and controllers wired to an in-memory Graph + DB behind a seeded chaos layer.
 * Only true externals are faked: axios (Microsoft Graph), the TypeORM repositories, and the
 * token services. Everything in between runs the production code paths.
 */
export function buildChaosWorld(mockedAxios: AxiosMockLike, options: WorldOptions) {
  const metrics = new ChaosMetrics();
  const engine = new ChaosEngine(options.seed, options.graphRates ?? {}, options.graphLatencyMs ?? { min: 0, max: 0 });
  const graph = new ChaosGraph(engine, metrics);
  const db = new ChaosDb(engine, metrics, options.dbLatencyMs ?? { min: 0, max: 0 });
  graph.install(mockedAxios);

  // ── seed tenants ────────────────────────────────────────────────────
  const tenantByKey = new Map<string, MicrosoftTenant>();
  const tenantSeeds = options.tenants ?? [{ key: 'T1', tenantId: 'tenant-t1-guid' }];
  for (const seed of tenantSeeds) {
    const tenant = db.addTenant({
      tenantId: seed.tenantId,
      status: seed.status ?? MicrosoftTenantStatus.ACTIVE,
      isActive: seed.isActive ?? true,
    });
    tenantByKey.set(seed.key, tenant);
  }
  const defaultTenant = tenantByKey.get(tenantSeeds[0].key);
  if (!defaultTenant) throw new Error('chaos world: no default tenant');

  // ── seed users + subscriptions ──────────────────────────────────────
  const usersByExternalId = new Map<string, MicrosoftUser>();
  let seedSubSeq = 0;
  for (const seed of options.users ?? []) {
    const tenant = tenantByKey.get(seed.tenantKey ?? tenantSeeds[0].key);
    if (!tenant) throw new Error(`chaos world: unknown tenantKey for ${seed.externalUserId}`);
    const msUserId = `ms-${seed.externalUserId}`;

    if (seed.inGraphDirectory !== false) {
      graph.seedUser(seed.email, msUserId);
    }
    if (seed.inDb === false) continue;

    const appOnly = seed.kind === 'app-only' || seed.kind === 'dual';
    const delegated = seed.kind === 'delegated' || seed.kind === 'dual';
    const user = db.addUser({
      externalUserId: seed.externalUserId,
      isActive: seed.isActive ?? true,
      status: seed.status ?? MicrosoftUserStatus.ACTIVE,
      refreshToken: delegated ? `refresh-${seed.externalUserId}` : null,
      accessToken: delegated ? `access-${seed.externalUserId}` : null,
      tenant: appOnly ? tenant : null,
      microsoftUserId: appOnly ? msUserId : null,
      userPrincipalName: appOnly ? seed.email : null,
    });
    usersByExternalId.set(seed.externalUserId, user);

    if (seed.sub) {
      seedSubSeq += 1;
      const subscriptionId = `seed-sub-${seedSubSeq}`;
      const isAppOnlySub = seed.sub.mode === 'app-only';
      const resource = isAppOnlySub ? `/users/${msUserId}/events` : '/me/events';
      const now = Date.now();
      db.addSubscription({
        subscriptionId,
        userId: user.id,
        tenantId: isAppOnlySub ? tenant.tenantId : null,
        microsoftUserId: isAppOnlySub ? msUserId : null,
        resource,
        changeType: 'created,updated,deleted',
        clientState: isAppOnlySub
          ? `tenant_${tenant.tenantId}_user_${user.id}_seed`
          : `user_${user.id}_seed`,
        notificationUrl: 'https://host.example.com/api/calendar/webhook',
        expirationDateTime: new Date(seed.sub.expired ? now - 1 * HOUR_MS : now + 48 * HOUR_MS),
        lastNotificationAt: new Date(seed.sub.stale ? now - 72 * HOUR_MS : now - 1 * HOUR_MS),
        createdAt: new Date(now - 96 * HOUR_MS),
      });
      if (seed.sub.presentAtGraph !== false) {
        graph.seedSubscription({
          id: subscriptionId,
          resource,
          changeType: 'created,updated,deleted',
          clientState: 'seeded',
          notificationUrl: 'https://host.example.com/api/calendar/webhook',
        });
      }
    }
  }

  // ── fakes for the remaining externals ───────────────────────────────
  const events: CapturedEvent[] = [];
  const eventEmitter = {
    emit: (event: string, payload: unknown): boolean => {
      events.push({ event, payload });
      return true;
    },
  } as unknown as EventEmitter2;

  const microsoftAuthService = {
    getUserAccessToken: async (params: { internalUserId?: number }): Promise<string> => {
      const key = String(params.internalUserId ?? '?');
      const injected = engine.decide('auth.delegatedToken', key);
      if (injected) {
        metrics.recordInjected('auth.delegatedToken', injected.response?.status ?? 'network');
        throw new Error(`chaos: delegated token unavailable for user ${key}`);
      }
      return `delegated-token-${key}`;
    },
  } as unknown as MicrosoftAuthService;

  const appOnlyAuthService = {
    getAccessToken: async (tenantId: string): Promise<string> => {
      const injected = engine.decide('auth.appOnlyToken', tenantId);
      if (injected) {
        metrics.recordInjected('auth.appOnlyToken', injected.response?.status ?? 'network');
        throw new Error(`chaos: app-only token unavailable for tenant ${tenantId}`);
      }
      return `app-only-token-${tenantId}`;
    },
    getTenantId: (): string => defaultTenant.tenantId,
    invalidateCache: (_tenantId?: string): void => {
      metrics.mark('auth:invalidateCache');
    },
    isEnabled: (): boolean => true,
  } as unknown as AppOnlyAuthService;

  const userIdConverter = {
    externalToInternal: async (externalUserId: string): Promise<number> => {
      metrics.recordDb('converter.externalToInternal');
      const user = db.users.find((u) => u.externalUserId === externalUserId);
      if (!user) throw new Error(`chaos: no microsoft_users row for ${externalUserId}`);
      return user.id;
    },
    internalToExternal: async (internalUserId: number): Promise<string> => {
      metrics.recordDb('converter.internalToExternal');
      const user = db.users.find((u) => u.id === internalUserId);
      if (!user) throw new Error(`chaos: no microsoft_users row for id ${internalUserId}`);
      return user.externalUserId;
    },
  } as unknown as UserIdConverterService;

  const rateLimiter = {
    acquirePermit: async (): Promise<void> => undefined,
    recordSuccess: async (): Promise<void> => undefined,
    record503Failure: async (): Promise<void> => undefined,
    handleRateLimitResponse: async (): Promise<void> => undefined,
  } as unknown as GraphRateLimiterService;

  const config: MicrosoftOutlookConfig = {
    clientId: 'chaos-client-id',
    clientSecret: 'chaos-client-secret',
    redirectPath: 'auth/microsoft/callback',
    backendBaseUrl: 'https://host.example.com',
    basePath: 'api',
    calendarWebhookPath: '/calendar/webhook',
  };

  const subscriptionRepo = db.buildSubscriptionRepo() as unknown as OutlookWebhookSubscriptionRepository;
  const userOrmRepo = db.buildUserOrmRepo() as unknown as Repository<MicrosoftUser>;
  const tenantOrmRepo = db.buildTenantOrmRepo() as unknown as Repository<MicrosoftTenant>;
  const tenantConnectionRepo = db.buildTenantConnectionRepo() as unknown as MicrosoftTenantRepository;

  // ── REAL services under test ────────────────────────────────────────
  const subscriptionService = new MicrosoftSubscriptionService(
    microsoftAuthService,
    appOnlyAuthService,
    subscriptionRepo,
    eventEmitter,
    config,
    userOrmRepo,
    userIdConverter,
    rateLimiter,
  );
  const tenantUserService = new TenantUserService(tenantOrmRepo, userOrmRepo, appOnlyAuthService, config);
  const provisioningService = new TenantProvisioningService(
    tenantUserService,
    subscriptionService,
    tenantConnectionRepo,
    eventEmitter,
  );
  const healthService = new HealthService(tenantUserService, subscriptionService, subscriptionRepo, eventEmitter);

  const tenantAuthController = new TenantAuthController(
    appOnlyAuthService,
    tenantConnectionRepo,
    provisioningService,
    tenantUserService,
    subscriptionService,
  );
  const healthController = new HealthController(healthService);

  return {
    metrics,
    engine,
    graph,
    db,
    events,
    tenants: tenantByKey,
    defaultTenantId: defaultTenant.tenantId,
    services: { subscriptionService, tenantUserService, provisioningService, healthService },
    controllers: { tenantAuthController, healthController },
    helpers: {
      msIdOf: (externalUserId: string): string => `ms-${externalUserId}`,
      internalIdOf: (externalUserId: string): number => {
        const user = usersByExternalId.get(externalUserId);
        if (!user) throw new Error(`chaos world: unseeded user ${externalUserId}`);
        return user.id;
      },
      activeDbSubsOf: (externalUserId: string) => {
        const user = db.users.find((u) => u.externalUserId === externalUserId);
        return user ? db.activeSubsOfUser(user.id) : [];
      },
      eventsNamed: (event: string): CapturedEvent[] => events.filter((e) => e.event === event),
    },
  };
}

export type ChaosWorld = ReturnType<typeof buildChaosWorld>;
