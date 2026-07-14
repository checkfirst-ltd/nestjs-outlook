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

jest.mock('axios');
const mockedAxios = axios as jest.Mocked<typeof axios> as unknown as AxiosMockLike;

jest.setTimeout(120_000);

/**
 * Chaos + scale tests for the tenant disconnect/purge flow: batched subscription deletion at
 * Graph, bulk local deactivation, user-mapping teardown with token revocation, and the
 * 202-background controller contract — all against a disrupted Graph/DB.
 */
describe('Tenant disconnect (purge) — chaos at scale', () => {
  const SEED = Number(process.env.CHAOS_SEED ?? 20260714);
  const TENANT = 'tenant-t1-guid';

  /** 450 app-only-only users + 150 dual (delegated tokens too), each with one app-only sub. */
  const seedTenantUsers = (): SeedUser[] => [
    ...Array.from({ length: 450 }, (_, i) => ({
      externalUserId: `ao-${i}`,
      email: `ao-${i}@contoso.com`,
      kind: 'app-only' as const,
      sub: { mode: 'app-only' as const },
    })),
    ...Array.from({ length: 150 }, (_, i) => ({
      externalUserId: `dual-${i}`,
      email: `dual-${i}@contoso.com`,
      kind: 'dual' as const,
      sub: { mode: 'app-only' as const },
    })),
  ];

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

  it('baseline: 600 subscriptions deleted via $batch (30 calls, not 600) + ONE bulk deactivate', async () => {
    const world = buildChaosWorld(mockedAxios, { seed: SEED, users: seedTenantUsers() });
    const svc = world.services.subscriptionService;

    const result = await drain(svc.deleteAllAppOnlySubscriptionsForTenant(TENANT));

    expect(result).toMatchObject({ totalFound: 600, successfullyDeleted: 600, failedToDelete: 0 });
    // Consumption: ceil(600/20) = 30 batch calls; zero per-subscription DELETEs.
    expect(world.metrics.attemptsFor('batch')).toBe(30);
    expect(world.metrics.attemptsFor('subs.delete')).toBe(0);
    // Local teardown is ONE set-based statement, not 600 row updates.
    expect(world.metrics.dbCallsFor('subs.deactivateAllByTenantId')).toBe(1);
    expect(world.metrics.dbCallsFor('subs.deactivate')).toBe(0);

    expect(world.db.subscriptions.filter((s) => s.tenantId === TENANT && s.isActive)).toHaveLength(0);
    expect(world.graph.subscriptionIdsForResourcePrefix('/users/')).toHaveLength(0);

    console.log(world.metrics.report(`purge baseline 600 subs seed=${SEED}`));
  });

  it('batch chaos: outer 429 retries + per-item failures — exact tallies, chunking stable, locals always deactivated', async () => {
    const users = seedTenantUsers();
    const world = buildChaosWorld(mockedAxios, { seed: SEED, users });
    const svc = world.services.subscriptionService;

    // First 3 whole-batch attempts are throttled (deterministic retries), then weather clears.
    world.engine.failTimes('batch', '*', 3, 429);
    // 50 subscriptions are genuinely "already gone" at Graph (drift → inner 404 = success),
    // 30 hard-fail per item.
    const dbSubs = world.db.subscriptions.filter((s) => s.tenantId === TENANT);
    const goneIds = dbSubs.slice(0, 50).map((s) => s.subscriptionId);
    const brokenIds = dbSubs.slice(50, 80).map((s) => s.subscriptionId);
    for (const id of goneIds) world.graph.subscriptions.delete(id);
    for (const id of brokenIds) world.engine.alwaysFail('batch.item', id, 500);

    const result = await drain(svc.deleteAllAppOnlySubscriptionsForTenant(TENANT));

    // 404s count as deleted; only the 30 hard failures are reported.
    expect(result.successfullyDeleted).toBe(570);
    expect(result.failedToDelete).toBe(30);
    expect(result.successfullyDeleted + result.failedToDelete).toBe(result.totalFound);
    expect(result.errors).toHaveLength(30);
    expect(new Set(result.errors.map((e) => e.subscriptionId))).toEqual(new Set(brokenIds));
    // Batch-response correlation: the reported deleted ids are exactly the non-broken set.
    expect(result.deletedSubscriptionIds).toHaveLength(570);
    const broken = new Set(brokenIds);
    expect(result.deletedSubscriptionIds.every((id) => !broken.has(id))).toBe(true);
    // Chunking stayed at 30 logical batches; the outer retries added exactly 3 attempts.
    expect(world.metrics.attemptsFor('batch')).toBe(33);
    // Local rows are ALL deactivated regardless of the per-item Graph outcome.
    expect(world.db.subscriptions.filter((s) => s.tenantId === TENANT && s.isActive)).toHaveLength(0);
    // The 30 broken ones remain at Graph (they truly were not deleted).
    expect(world.graph.subscriptionIdsForResourcePrefix('/users/')).toHaveLength(30);

    console.log(world.metrics.report(`purge batch-chaos seed=${SEED}`));
  });

  it('mapping teardown (default): dual rows unmapped keeping delegated login, app-only rows deleted — zero revocations', async () => {
    const world = buildChaosWorld(mockedAxios, { seed: SEED, users: seedTenantUsers() });

    const result = await drain(world.services.tenantUserService.clearTenantUserMappings(TENANT));

    expect(result).toMatchObject({
      delegatedRowsUnmapped: 150,
      appOnlyRowsDeleted: 450,
      tokensRevoked: 0,
      tokenRevocationFailures: 0,
    });
    expect(world.metrics.attemptsFor('auth.revoke')).toBe(0);
    // Consumption: the teardown is exactly one bulk UPDATE + one bulk DELETE, not 600 row ops.
    expect(world.metrics.dbCallsFor('users.qb.update')).toBe(1);
    expect(world.metrics.dbCallsFor('users.qb.delete')).toBe(1);
    expect(world.metrics.dbCallsFor('users.save')).toBe(0);
    // Dual users survive with delegated login intact and the app-only mapping stripped.
    const dualRows = world.db.users.filter((u) => u.externalUserId.startsWith('dual-'));
    expect(dualRows).toHaveLength(150);
    for (const row of dualRows) {
      expect(row.refreshToken).not.toBeNull();
      expect(row.tenant).toBeNull();
      expect(row.microsoftUserId).toBeNull();
    }
    expect(world.db.users.filter((u) => u.externalUserId.startsWith('ao-'))).toHaveLength(0);
  });

  it('revocation chaos: 150 tokens revoked single-shot at bounded concurrency; failures counted, teardown never aborts', async () => {
    const world = buildChaosWorld(mockedAxios, {
      seed: SEED,
      users: seedTenantUsers(),
      graphLatencyMs: { min: 1, max: 10 },
    });

    // 20 delegated tokens hard-fail at the logout endpoint.
    const failingTokens = Array.from({ length: 20 }, (_, i) => `refresh-dual-${i}`);
    for (const token of failingTokens) world.engine.alwaysFail('auth.revoke', token, 500);

    const result = await drain(
      world.services.tenantUserService.clearTenantUserMappings(TENANT, { revokeDelegatedTokens: true }),
    );

    expect(result.tokensRevoked).toBe(130);
    expect(result.tokenRevocationFailures).toBe(20);
    // Revocation is best-effort single-shot: exactly one attempt per token, no retry storm.
    expect(world.metrics.attemptsFor('auth.revoke')).toBe(150);
    for (const token of failingTokens) {
      expect(world.metrics.attemptsForKey('auth.revoke', token)).toBe(1);
    }
    // Bounded concurrency at the revocation endpoint — parallel (>1) but capped at 5.
    expect(world.metrics.peakInFlight).toBeLessThanOrEqual(5);
    expect(world.metrics.peakInFlight).toBeGreaterThan(1);
    // Aggressive path removes every tenant row despite the failures.
    expect(result.appOnlyRowsDeleted).toBe(600);
    expect(world.db.users).toHaveLength(0);

    console.log(world.metrics.report(`revocation chaos seed=${SEED}`));
  });

  it('token unavailable: falls back to local-only deactivation with zero Graph traffic', async () => {
    const world = buildChaosWorld(mockedAxios, { seed: SEED, users: seedTenantUsers() });
    world.engine.alwaysFail('auth.appOnlyToken', TENANT, 500);

    const result = await drain(world.services.subscriptionService.deleteAllAppOnlySubscriptionsForTenant(TENANT));

    expect(result.localOnlyDeactivated).toBe(600);
    expect(result.successfullyDeleted).toBe(0);
    expect(world.metrics.attemptsFor('batch')).toBe(0);
    expect(world.db.subscriptions.filter((s) => s.tenantId === TENANT && s.isActive)).toHaveLength(0);
  });

  it('host contract: soft disconnect (no purge) deactivates synchronously and touches NOTHING else', async () => {
    const users = seedTenantUsers().slice(0, 80);
    const world = buildChaosWorld(mockedAxios, { seed: SEED, users });

    const response = await drain(world.controllers.tenantAuthController.disconnect(TENANT));

    expect(response.message).toContain('disconnected successfully');
    // Synchronous: the tenant is already inactive at response time.
    expect(world.db.tenants.find((t) => t.tenantId === TENANT)?.isActive).toBe(false);
    // Soft = zero Graph traffic, subscriptions and user mappings left fully intact.
    expect(world.metrics.attempts.size).toBe(0);
    expect(world.db.subscriptions.filter((s) => s.tenantId === TENANT && s.isActive)).toHaveLength(80);
    expect(world.db.users.every((u) => u.tenant !== null)).toBe(true);
  });

  it('host contract: DELETE connection?purge=true returns 202-style before teardown; tenant deactivated LAST', async () => {
    const users = seedTenantUsers().slice(0, 120);
    const world = buildChaosWorld(mockedAxios, {
      seed: SEED,
      users,
      graphLatencyMs: { min: 2, max: 15 },
    });

    const response = await world.controllers.tenantAuthController.disconnect(TENANT, 'true', 'false');

    expect(response.message).toContain('background');
    // At response time the teardown has not run: the tenant connection is still active.
    const tenant = world.db.tenants.find((t) => t.tenantId === TENANT);
    expect(tenant?.isActive).toBe(true);

    await drainUntil(() => world.metrics.timeline.includes('tenant:deactivate'));

    // Ordering: every Graph batch delete happened BEFORE the tenant was deactivated
    // (the app-only token must stay valid during teardown), and the cache was dropped after.
    // Guard against vacuity first: the purge really did delete at Graph.
    const lastBatchAt = world.metrics.lastIndexOf('graph:batch');
    expect(lastBatchAt).toBeGreaterThan(-1);
    const deactivateAt = world.metrics.timeline.indexOf('tenant:deactivate');
    expect(lastBatchAt).toBeLessThan(deactivateAt);
    expect(world.metrics.timeline.indexOf('auth:invalidateCache')).toBeGreaterThan(deactivateAt);

    expect(tenant?.isActive).toBe(false);
    expect(world.db.subscriptions.filter((s) => s.tenantId === TENANT && s.isActive)).toHaveLength(0);
    // Mapping teardown ran too: no row still points at the tenant.
    expect(world.db.users.every((u) => u.tenant === null)).toBe(true);
  });
});
