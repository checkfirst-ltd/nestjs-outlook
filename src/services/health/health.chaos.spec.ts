import { Logger } from '@nestjs/common';
import axios from 'axios';
import {
  buildChaosWorld,
  SeedUser,
  AxiosMockLike,
  installChaosTimers,
  uninstallChaosTimers,
  drain,
  drainUntil,
} from '../../mock/chaos';
import { UserHealth, HealthCheckResult } from './health.service';
import { UserHealthStatus } from '../../enums/user-health-status.enum';
import { MicrosoftUserStatus } from '../../enums/microsoft-user-status.enum';
import { MicrosoftTenantStatus } from '../../enums/microsoft-tenant-status.enum';
import { OutlookEventTypes } from '../../enums/event-types.enum';

jest.mock('axios');
const mockedAxios = axios as jest.Mocked<typeof axios> as unknown as AxiosMockLike;

jest.setTimeout(120_000);

/**
 * Chaos + scale tests for HealthService.checkUsers / recoverUsers over a 400-user population in
 * every reachable state, across two tenants (one ACTIVE, one CONSENT_REVOKED), with Graph
 * disruption on the verify and create paths.
 *
 * Behavioural note surfaced by the realistic repo fakes: `checkUsers` reads subscriptions via
 * `findActiveByUserIds`, whose REAL query excludes rows past `expirationDateTime` — so an
 * expired subscription is diagnosed as NO_SUBSCRIPTION (not SUBSCRIPTION_EXPIRED). Both are
 * recoverable, so recovery behaviour is identical; the assertions below encode the real verdict.
 */
describe('HealthService — chaos at scale', () => {
  const SEED = Number(process.env.CHAOS_SEED ?? 20260714);
  const T1 = 'tenant-t1-guid';
  const T2 = 'tenant-t2-guid';

  const group = (
    n: number,
    prefix: string,
    build: (i: number) => Omit<SeedUser, 'externalUserId' | 'email'>,
  ): SeedUser[] =>
    Array.from({ length: n }, (_, i) => ({
      externalUserId: `${prefix}-${i}`,
      email: `${prefix}-${i}@contoso.com`,
      ...build(i),
    }));

  /** 400 users covering every verdict. */
  const seedPopulation = (): SeedUser[] => [
    ...group(150, 'healthy-ao', () => ({ kind: 'app-only', sub: { mode: 'app-only' } })),
    ...group(40, 'healthy-del', () => ({ kind: 'delegated', sub: { mode: 'delegated' } })),
    ...group(50, 'nosub', () => ({ kind: 'app-only', sub: null })),
    ...group(30, 'expired', () => ({ kind: 'app-only', sub: { mode: 'app-only', expired: true } })),
    ...group(30, 'stale-ao', () => ({ kind: 'app-only', sub: { mode: 'app-only', stale: true } })),
    ...group(10, 'stale-del', () => ({ kind: 'delegated', sub: { mode: 'delegated', stale: true } })),
    ...group(25, 'missing', () => ({ kind: 'app-only', sub: { mode: 'app-only', presentAtGraph: false } })),
    ...group(20, 'corrupted', () => ({
      kind: 'delegated',
      status: MicrosoftUserStatus.CORRUPTED,
      sub: { mode: 'delegated' },
    })),
    ...group(15, 'revoked', () => ({ kind: 'app-only', tenantKey: 'T2', sub: null })),
    ...group(10, 'inactive', () => ({ kind: 'app-only', isActive: false, sub: null })),
    ...group(8, 'bare', () => ({ kind: 'bare' })),
    ...group(12, 'ghost', () => ({ kind: 'bare', inDb: false })),
  ];

  const TENANTS = [
    { key: 'T1', tenantId: T1 },
    { key: 'T2', tenantId: T2, status: MicrosoftTenantStatus.CONSENT_REVOKED },
  ];

  const idsOf = (seeds: SeedUser[]): string[] => seeds.map((s) => s.externalUserId);

  const histogram = (healths: UserHealth[]): Record<string, number> => {
    const counts: Record<string, number> = {};
    for (const h of healths) counts[h.status] = (counts[h.status] ?? 0) + 1;
    return counts;
  };

  beforeAll(() => {
    Logger.overrideLogger(false);
  });

  beforeEach(() => {
    jest.clearAllMocks();
    installChaosTimers();
  });

  afterEach(() => {
    uninstallChaosTimers();
  });

  it('DB-only check of 400 mixed-state users: exact verdict histogram, ZERO Graph traffic, two bulk reads', async () => {
    const seeds = seedPopulation();
    const world = buildChaosWorld(mockedAxios, { seed: SEED, tenants: TENANTS, users: seeds });

    const healths = await drain(world.services.healthService.checkUsers(idsOf(seeds)));

    expect(healths).toHaveLength(400);
    expect(histogram(healths)).toEqual({
      [UserHealthStatus.HEALTHY]: 215, // 150 + 40 + 25 drifted (verify off ⇒ DB-healthy)
      [UserHealthStatus.NO_SUBSCRIPTION]: 80, // 50 without + 30 expired (real query excludes expired)
      [UserHealthStatus.SUBSCRIPTION_STALE]: 40,
      [UserHealthStatus.NEEDS_REAUTH]: 20,
      [UserHealthStatus.NEEDS_ADMIN]: 15,
      [UserHealthStatus.INACTIVE]: 10,
      [UserHealthStatus.NOT_MAPPED]: 8,
      [UserHealthStatus.UNKNOWN]: 12,
    });

    // Consumption: a DB-only check generates ZERO Graph traffic on ANY route…
    expect(world.metrics.attempts.size).toBe(0);
    // …and reads the database in TWO bulk queries, not 400.
    expect(world.metrics.dbCallsFor('users.find')).toBe(1);
    expect(world.metrics.dbCallsFor('subs.findActiveByUserIds')).toBe(1);

    // Input order is preserved (the documented contract) — even through the bulk path.
    const ids = idsOf(seeds);
    for (let i = 0; i < ids.length; i++) {
      expect(healths[i].externalUserId).toBe(ids[i]);
    }

    console.log(world.metrics.report(`health check DB-only N=400 seed=${SEED}`));
  });

  it('verifyAtGraph: drifted subscriptions detected exactly; inconclusive Graph answers never downgrade a verdict', async () => {
    const seeds = seedPopulation();
    const world = buildChaosWorld(mockedAxios, { seed: SEED, tenants: TENANTS, users: seeds });

    // 6 genuinely-healthy users whose verify call hard-fails (inconclusive 'unknown').
    const inconclusive = seeds.filter((s) => s.externalUserId.startsWith('healthy-ao')).slice(0, 6);
    for (const s of inconclusive) {
      const subId = world.helpers.activeDbSubsOf(s.externalUserId)[0].subscriptionId;
      world.engine.alwaysFail('subs.get', subId, 500);
    }

    const healths = await drain(world.services.healthService.checkUsers(idsOf(seeds), { verifyAtGraph: true }));

    const counts = histogram(healths);
    expect(counts[UserHealthStatus.MISSING_AT_GRAPH]).toBe(25);
    expect(counts[UserHealthStatus.HEALTHY]).toBe(190); // 215 − 25 drifted; the 6 inconclusive stay HEALTHY
    // Only DB-healthy users were verified: 215 logical calls; the 6 sticky-500s retried
    // (maxRetries=3 ⇒ 4 attempts each), everything else answered once.
    expect(world.metrics.attemptsFor('subs.get')).toBe(209 + 6 * 4);
    for (const s of inconclusive) {
      const health = healths.find((h) => h.externalUserId === s.externalUserId);
      expect(health?.status).toBe(UserHealthStatus.HEALTHY);
    }
  });

  it('verify storm: random 429s + latency — DB-decided verdicts stay exact, totals conserve, ceiling holds', async () => {
    const seeds = seedPopulation();
    const world = buildChaosWorld(mockedAxios, {
      seed: SEED,
      tenants: TENANTS,
      users: seeds,
      graphRates: { throttle429: 0.15 },
      graphLatencyMs: { min: 1, max: 10 },
    });

    const healths = await drain(world.services.healthService.checkUsers(idsOf(seeds), { verifyAtGraph: true }));

    const counts = histogram(healths);
    // Conservation across all verdicts.
    expect(healths).toHaveLength(400);
    // Verdicts that never touch Graph are immune to Graph weather.
    expect(counts[UserHealthStatus.NEEDS_REAUTH]).toBe(20);
    expect(counts[UserHealthStatus.NEEDS_ADMIN]).toBe(15);
    expect(counts[UserHealthStatus.INACTIVE]).toBe(10);
    expect(counts[UserHealthStatus.NOT_MAPPED]).toBe(8);
    expect(counts[UserHealthStatus.UNKNOWN]).toBe(12);
    expect(counts[UserHealthStatus.NO_SUBSCRIPTION]).toBe(80);
    expect(counts[UserHealthStatus.SUBSCRIPTION_STALE]).toBe(40);
    // Graph-decided split can shift under throttling exhaustion, but only between these two.
    expect((counts[UserHealthStatus.HEALTHY] ?? 0) + (counts[UserHealthStatus.MISSING_AT_GRAPH] ?? 0)).toBe(215);
    expect(counts[UserHealthStatus.MISSING_AT_GRAPH] ?? 0).toBeLessThanOrEqual(25);
    expect(world.metrics.totalInjected()).toBeGreaterThan(0);
    expect(world.metrics.peakInFlight).toBeLessThanOrEqual(5);

    console.log(world.metrics.report(`health verify storm N=400 seed=${SEED}`));
  });

  it('recoverUsers: fixes every recoverable state via the right auth mode, reports the rest, and converges on rerun', async () => {
    const seeds = seedPopulation();
    const world = buildChaosWorld(mockedAxios, { seed: SEED, tenants: TENANTS, users: seeds });

    // Planned create failures: 5 app-only recoveries and 4 delegated recoveries hard-fail.
    const appOnlyFails = seeds.filter((s) => s.externalUserId.startsWith('nosub')).slice(0, 5);
    const delegatedFails = seeds.filter((s) => s.externalUserId.startsWith('stale-del')).slice(0, 4);
    for (const s of appOnlyFails) world.engine.alwaysFail('subs.create', world.helpers.msIdOf(s.externalUserId), 500);
    for (const s of delegatedFails) {
      world.engine.alwaysFail('subs.create', `me:${world.helpers.internalIdOf(s.externalUserId)}`, 500);
    }

    const result = await drain(world.services.healthService.recoverUsers(idsOf(seeds), { verifyAtGraph: true }));

    // Recoverable = 50 nosub + 30 expired + 30 stale-ao + 10 stale-del + 25 drifted = 145.
    expect(result.total).toBe(400);
    expect(result.healthy).toBe(190);
    expect(result.recovered).toBe(145 - 9);
    expect(result.failed).toBe(9);
    expect(result.unrecoverable).toBe(20 + 15 + 10 + 8 + 12);

    // Routing: delegated users were recovered via /me/events, app-only via /users/{id}/events.
    const delegatedRecovered = result.results.filter(
      (r) => r.externalUserId.startsWith('stale-del') && r.action === 'recreated',
    );
    expect(delegatedRecovered).toHaveLength(6);
    for (const r of result.results.filter((x) => x.action === 'recreated')) {
      const subs = world.helpers.activeDbSubsOf(r.externalUserId);
      expect(subs).toHaveLength(1);
      const expectDelegated = r.externalUserId.startsWith('stale-del');
      expect(subs[0].resource.startsWith(expectDelegated ? '/me/' : '/users/')).toBe(true);
    }
    // Unrecoverable users were REPORTED, never re-created. Check BOTH create-key shapes
    // (app-only keys by Microsoft id, delegated keys by `me:{internalId}`) so the assertion
    // can't be satisfied vacuously by looking up the wrong namespace.
    for (const prefix of ['corrupted', 'revoked', 'inactive', 'bare'] as const) {
      const samples = seeds.filter((s) => s.externalUserId.startsWith(prefix));
      expect(samples.length).toBeGreaterThan(0);
      for (const sample of samples) {
        expect(world.metrics.attemptsForKey('subs.create', world.helpers.msIdOf(sample.externalUserId))).toBe(0);
        expect(
          world.metrics.attemptsForKey('subs.create', `me:${world.helpers.internalIdOf(sample.externalUserId)}`),
        ).toBe(0);
      }
    }
    // Belt-and-braces: total create attempts equal exactly the recovered creates plus the
    // exhausted retries of the 9 planned failures (maxRetries=7 ⇒ 8 attempts each).
    expect(world.metrics.attemptsFor('subs.create')).toBe(136 + 9 * 8);
    const completed = world.helpers.eventsNamed(OutlookEventTypes.USER_HEALTH_RECOVERY_COMPLETED);
    expect(completed).toHaveLength(1);

    // Convergence: a second run finds the recovered users HEALTHY and re-attempts only the 9 failures.
    const rerun = await drain(world.services.healthService.recoverUsers(idsOf(seeds), { verifyAtGraph: true }));
    expect(rerun.healthy).toBe(190 + 136);
    expect(rerun.recovered).toBe(0);
    expect(rerun.failed).toBe(9);
    expect(rerun.unrecoverable).toBe(65);

    console.log(world.metrics.report(`health recover N=400 (two runs) seed=${SEED}`));
  });

  it('host contract: POST health/recover returns before the background recovery completes (mixed population)', async () => {
    const seeds: SeedUser[] = [
      ...group(50, 'ok', () => ({ kind: 'app-only' as const, sub: { mode: 'app-only' as const } })),
      ...group(40, 'fix', () => ({ kind: 'app-only' as const, sub: null })),
      ...group(20, 'stuck', () => ({
        kind: 'delegated' as const,
        status: MicrosoftUserStatus.CORRUPTED,
        sub: null,
      })),
    ];
    const world = buildChaosWorld(mockedAxios, {
      seed: SEED,
      tenants: TENANTS,
      users: seeds,
      graphLatencyMs: { min: 2, max: 15 },
    });

    const response = world.controllers.healthController.recoverUsers({ externalUserIds: idsOf(seeds) });

    expect(response.totalRequested).toBe(110);
    expect(world.helpers.eventsNamed(OutlookEventTypes.USER_HEALTH_RECOVERY_COMPLETED)).toHaveLength(0);

    await drainUntil(
      () => world.helpers.eventsNamed(OutlookEventTypes.USER_HEALTH_RECOVERY_COMPLETED).length === 1,
    );
    const summary = world.helpers.eventsNamed(OutlookEventTypes.USER_HEALTH_RECOVERY_COMPLETED)[0]
      .payload as HealthCheckResult;
    // Latency-only chaos ⇒ deterministic outcome: the exact split, not just the sum.
    expect(summary).toMatchObject({ total: 110, healthy: 50, recovered: 40, unrecoverable: 20, failed: 0 });
    for (const s of seeds.filter((x) => x.externalUserId.startsWith('fix'))) {
      expect(world.helpers.activeDbSubsOf(s.externalUserId)).toHaveLength(1);
    }
  });

  it('single-user endpoint passthrough: unknown id yields UNKNOWN', async () => {
    const world = buildChaosWorld(mockedAxios, { seed: SEED, tenants: TENANTS, users: [] });

    const health = await drain(world.controllers.healthController.getUserHealth('nobody', 'false'));

    expect(health.status).toBe(UserHealthStatus.UNKNOWN);
    expect(health.connected).toBe(false);
  });

  it('GET health/:id with verifyAtGraph under throttling: a real healthy user resolves HEALTHY through retries', async () => {
    const seeds: SeedUser[] = [
      { externalUserId: 'one', email: 'one@contoso.com', kind: 'app-only', sub: { mode: 'app-only' } },
    ];
    const world = buildChaosWorld(mockedAxios, { seed: SEED, tenants: TENANTS, users: seeds });
    const subId = world.helpers.activeDbSubsOf('one')[0].subscriptionId;
    world.engine.failTimes('subs.get', subId, 2, 429); // deterministic transient throttling

    const health = await drain(world.controllers.healthController.getUserHealth('one', 'true'));

    expect(health.status).toBe(UserHealthStatus.HEALTHY);
    expect(health.connected).toBe(true);
    expect(world.metrics.attemptsForKey('subs.get', subId)).toBe(3); // 2 throttles + 1 success
  });

  it('POST health/check endpoint: bulk verdicts through the controller (DTO path), order preserved', async () => {
    const seeds: SeedUser[] = [
      { externalUserId: 'ok-1', email: 'ok-1@contoso.com', kind: 'app-only', sub: { mode: 'app-only' } },
      { externalUserId: 'gone-1', email: 'gone-1@contoso.com', kind: 'app-only', sub: null },
      { externalUserId: 'dead-1', email: 'dead-1@contoso.com', kind: 'delegated', status: MicrosoftUserStatus.CORRUPTED, sub: null },
    ];
    const world = buildChaosWorld(mockedAxios, { seed: SEED, tenants: TENANTS, users: seeds });

    const healths = await drain(
      world.controllers.healthController.checkUsers({ externalUserIds: ['ok-1', 'gone-1', 'dead-1', 'ghost-1'] }),
    );

    expect(healths.map((h) => h.status)).toEqual([
      UserHealthStatus.HEALTHY,
      UserHealthStatus.NO_SUBSCRIPTION,
      UserHealthStatus.NEEDS_REAUTH,
      UserHealthStatus.UNKNOWN,
    ]);
  });

  it('recoverUser (single): one no-subscription user is recreated end to end', async () => {
    const seeds: SeedUser[] = [
      { externalUserId: 'solo', email: 'solo@contoso.com', kind: 'app-only', sub: null },
    ];
    const world = buildChaosWorld(mockedAxios, { seed: SEED, tenants: TENANTS, users: seeds });

    const report = await drain(world.services.healthService.recoverUser('solo'));

    expect(report.action).toBe('recreated');
    expect(report.status).toBe(UserHealthStatus.HEALTHY);
    expect(world.helpers.activeDbSubsOf('solo')).toHaveLength(1);
    expect(world.graph.subscriptionIdsForResourcePrefix('/users/')).toHaveLength(1);
  });
});
