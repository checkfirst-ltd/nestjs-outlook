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
import { BulkConnectResult } from './tenant-provisioning.service';
import { OutlookEventTypes } from '../../enums/event-types.enum';

jest.mock('axios');
const mockedAxios = axios as jest.Mocked<typeof axios> as unknown as AxiosMockLike;

jest.setTimeout(120_000);

/**
 * Chaos + scale tests for TenantProvisioningService.connectUsers — the REAL service stack
 * (provisioning → tenant-user → subscription) against an in-memory Graph/DB that injects
 * latency, throttling, server errors, and planned failures.
 */
describe('TenantProvisioningService — chaos at scale', () => {
  const SEED = Number(process.env.CHAOS_SEED ?? 20260714);
  const TENANT = 'tenant-t1-guid';

  const freshUsers = (n: number, prefix = 'fresh'): SeedUser[] =>
    Array.from({ length: n }, (_, i) => ({
      externalUserId: `${prefix}-${i}`,
      email: `${prefix}-${i}@contoso.com`,
      kind: 'bare' as const,
      inDb: false, // no microsoft_users row yet — the connect flow creates it
    }));

  const inputsOf = (seeds: SeedUser[]) =>
    seeds.map((s) => ({ externalUserId: s.externalUserId, email: s.email }));

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

  it('baseline (no chaos): connects 520 users with exact Graph/DB consumption (incl. >500 IN-chunking)', async () => {
    const seeds = freshUsers(520);
    const world = buildChaosWorld(mockedAxios, { seed: SEED, users: seeds });

    const result = await drain(world.services.provisioningService.connectUsers(TENANT, inputsOf(seeds)));

    expect(result).toMatchObject({ total: 520, connected: 520, skipped: 0, failed: 0 });

    // Graph consumption: exactly one lookup + one create per user, no dedupe deletes.
    expect(world.metrics.attemptsFor('users.lookup')).toBe(520);
    expect(world.metrics.attemptsFor('subs.create')).toBe(520);
    expect(world.metrics.attemptsFor('subs.delete')).toBe(0);
    // Bounded concurrency at the Graph boundary.
    expect(world.metrics.peakInFlight).toBeLessThanOrEqual(5);
    expect(world.metrics.peakInFlight).toBeGreaterThan(1);
    // The already-connected pre-filter is bulk queries, not N: 520 external ids exceed the
    // 500-per-IN chunk limit, so the user lookup runs exactly TWO chunked queries.
    expect(world.metrics.dbCallsFor('users.find')).toBe(2);
    expect(world.metrics.dbCallsFor('subs.findAllActiveByTenantId')).toBe(1);

    // End-state: every user holds exactly one active subscription, both locally and at Graph.
    for (const seed of seeds) {
      const subs = world.helpers.activeDbSubsOf(seed.externalUserId);
      expect(subs).toHaveLength(1);
      expect(world.graph.subscriptions.has(subs[0].subscriptionId)).toBe(true);
    }
    expect(world.graph.subscriptionIdsForResourcePrefix('/users/')).toHaveLength(520);

    const completed = world.helpers.eventsNamed(OutlookEventTypes.TENANT_USERS_BULK_CONNECT_COMPLETED);
    expect(completed).toHaveLength(1);
    expect(completed[0].payload as BulkConnectResult).toMatchObject({ connected: 520, failed: 0 });

    console.log(world.metrics.report(`bulk-connect baseline N=520 seed=${SEED}`));
  });

  it('chaos storm: 300 users under 429/503/500/network + latency — conserves tallies, no duplicates', async () => {
    const seeds = freshUsers(300, 'storm');
    const world = buildChaosWorld(mockedAxios, {
      seed: SEED,
      users: seeds,
      graphRates: { throttle429: 0.15, unavailable503: 0.05, serverError500: 0.05, networkError: 0.05 },
      graphLatencyMs: { min: 1, max: 20 },
    });

    const result = await drain(world.services.provisioningService.connectUsers(TENANT, inputsOf(seeds)));

    // Conservation: every requested user is accounted for exactly once.
    expect(result.total).toBe(300);
    expect(result.connected + result.failed).toBe(300);
    expect(result.skipped).toBe(0);
    expect(result.results).toHaveLength(300);
    // The storm actually happened; retry depth stays within the configured ceilings
    // (lookup maxRetries=3 ⇒ ≤4 attempts/user; create maxRetries=7 ⇒ ≤8 attempts/user).
    expect(world.metrics.totalInjected()).toBeGreaterThan(0);
    expect(world.metrics.attemptsFor('subs.create')).toBeGreaterThanOrEqual(result.connected);
    expect(world.metrics.attemptsFor('users.lookup')).toBeLessThanOrEqual(4 * 300);
    expect(world.metrics.attemptsFor('subs.create')).toBeLessThanOrEqual(8 * 300);
    // Concurrency ceiling holds under disruption.
    expect(world.metrics.peakInFlight).toBeLessThanOrEqual(5);

    // No duplicates and no orphans for FAIL-BEFORE-MUTATE disruption (this storm's mode):
    // connected ⇒ exactly one sub; failed ⇒ zero subs. The at-least-once case (response lost
    // AFTER Graph applied the create) is exercised separately below.
    for (const r of result.results) {
      const subs = world.helpers.activeDbSubsOf(r.externalUserId);
      if (r.success) {
        expect(subs).toHaveLength(1);
        expect(world.graph.subscriptions.has(subs[0].subscriptionId)).toBe(true);
      } else {
        expect(subs).toHaveLength(0);
        expect(r.error).toBeTruthy();
      }
    }
    expect(world.graph.subscriptionIdsForResourcePrefix('/users/')).toHaveLength(result.connected);

    // The completion event fires exactly once, with the same tallies the caller received.
    const completed = world.helpers.eventsNamed(OutlookEventTypes.TENANT_USERS_BULK_CONNECT_COMPLETED);
    expect(completed).toHaveLength(1);
    expect(completed[0].payload as BulkConnectResult).toMatchObject({
      connected: result.connected,
      failed: result.failed,
    });

    console.log(world.metrics.report(`bulk-connect storm N=300 seed=${SEED}`));
  });

  it('at-least-once reality: a create whose RESPONSE is lost after Graph applied it duplicates at Graph on retry', async () => {
    const seeds = freshUsers(40, 'lostack');
    const world = buildChaosWorld(mockedAxios, { seed: SEED, users: seeds });

    // For 10 users the first create attempt MUTATES Graph, then the response is lost
    // (network drop). Production retries the non-idempotent POST — the retry succeeds,
    // leaving TWO live Graph subscriptions while the local DB knows only the second.
    const lostAck = seeds.slice(0, 10);
    for (const s of lostAck) {
      world.engine.failTimesAfterExecute('subs.create', world.helpers.msIdOf(s.externalUserId), 1, 'network');
    }

    const result = await drain(world.services.provisioningService.connectUsers(TENANT, inputsOf(seeds)));

    expect(result).toMatchObject({ connected: 40, failed: 0 });
    for (const s of lostAck) {
      expect(world.helpers.activeDbSubsOf(s.externalUserId)).toHaveLength(1); // DB is clean…
    }
    // …but Graph holds an untracked duplicate per lost-ack user. This DOCUMENTS the current
    // production behaviour (known wart): both subs carry the SAME clientState, so the webhook
    // guard accepts notifications from the orphan too — duplicate notifications until it
    // expires (≤3 days). Remediation (create-time reconciliation) is a follow-up.
    expect(world.graph.subscriptionIdsForResourcePrefix('/users/')).toHaveLength(40 + 10);
  });

  it('orphan window: Graph create succeeds but the subscription SAVE fails — user fails cleanly, orphan documented', async () => {
    const seeds = freshUsers(30, 'orphan');
    const world = buildChaosWorld(mockedAxios, { seed: SEED, users: seeds });

    const dbFails = seeds.slice(0, 6);
    for (const s of dbFails) world.engine.alwaysFail('db.subs.save', s.externalUserId, 500);

    const result = await drain(world.services.provisioningService.connectUsers(TENANT, inputsOf(seeds)));

    expect(result.failed).toBe(6);
    expect(result.connected).toBe(24);
    for (const s of dbFails) {
      // The user is reported failed with no local subscription…
      expect(world.helpers.activeDbSubsOf(s.externalUserId)).toHaveLength(0);
    }
    // …but the Graph-side subscription was already created: an orphan per failed save.
    // Mitigation in production: its clientState was never persisted, so the webhook
    // clientState guard rejects the orphan's notifications, and it expires within 3 days.
    expect(world.graph.subscriptionIdsForResourcePrefix('/users/')).toHaveLength(24 + 6);
  });

  it('duplicate externalUserIds in one request are de-duplicated (no concurrent double-subscribe race)', async () => {
    const seeds = freshUsers(30, 'dup');
    const world = buildChaosWorld(mockedAxios, {
      seed: SEED,
      users: seeds,
      graphLatencyMs: { min: 5, max: 25 }, // overlap window for would-be concurrent duplicates
    });

    // Every user appears TWICE in the request.
    const doubled = [...inputsOf(seeds), ...inputsOf(seeds)];
    const result = await drain(world.services.provisioningService.connectUsers(TENANT, doubled));

    expect(result.total).toBe(30);
    expect(result.connected).toBe(30);
    expect(result.results).toHaveLength(30);
    // Exactly one create per user — the duplicate never raced the dedupe guard.
    expect(world.metrics.attemptsFor('subs.create')).toBe(30);
    for (const s of seeds) {
      expect(world.helpers.activeDbSubsOf(s.externalUserId)).toHaveLength(1);
    }
    expect(world.graph.subscriptionIdsForResourcePrefix('/users/')).toHaveLength(30);
  });

  it('planned failures are exact: 403 lookups fail fast (1 attempt), sticky 500 creates exhaust retries (8 attempts)', async () => {
    const seeds = freshUsers(200, 'plan');
    const world = buildChaosWorld(mockedAxios, { seed: SEED, users: seeds });

    const lookupFails = seeds.slice(0, 12);
    const createFails = seeds.slice(12, 22);
    for (const s of lookupFails) world.engine.alwaysFail('users.lookup', s.email, 403);
    for (const s of createFails) world.engine.alwaysFail('subs.create', world.helpers.msIdOf(s.externalUserId), 500);

    const result = await drain(world.services.provisioningService.connectUsers(TENANT, inputsOf(seeds)));

    expect(result.failed).toBe(22);
    expect(result.connected).toBe(178);

    // 403 is non-retryable: exactly one attempt, immediate failure.
    for (const s of lookupFails) {
      expect(world.metrics.attemptsForKey('users.lookup', s.email)).toBe(1);
    }
    // 500 is retryable: maxRetries=7 ⇒ exactly 8 attempts before giving up.
    for (const s of createFails) {
      expect(world.metrics.attemptsForKey('subs.create', world.helpers.msIdOf(s.externalUserId))).toBe(8);
    }
    // Failed creates never mutated Graph.
    expect(world.graph.subscriptionIdsForResourcePrefix('/users/')).toHaveLength(178);
  });

  it('transient 429s recover: fail-twice keys succeed on the 3rd attempt', async () => {
    const seeds = freshUsers(100, 'transient');
    const world = buildChaosWorld(mockedAxios, { seed: SEED, users: seeds });

    const flaky = seeds.slice(0, 20);
    for (const s of flaky) world.engine.failTimes('subs.create', world.helpers.msIdOf(s.externalUserId), 2, 429);

    const result = await drain(world.services.provisioningService.connectUsers(TENANT, inputsOf(seeds)));

    expect(result).toMatchObject({ connected: 100, failed: 0 });
    for (const s of flaky) {
      expect(world.metrics.attemptsForKey('subs.create', world.helpers.msIdOf(s.externalUserId))).toBe(3);
    }
    expect(world.metrics.injectedFor('subs.create', 429)).toBe(40);
  });

  it('existing delegated users: old /me/events subscription removed before app-only create — dedupe guard under disruption', async () => {
    const delegated: SeedUser[] = Array.from({ length: 40 }, (_, i) => ({
      externalUserId: `deleg-${i}`,
      email: `deleg-${i}@contoso.com`,
      kind: 'delegated',
      sub: { mode: 'delegated' },
    }));
    const fresh = freshUsers(110, 'new');
    const all = [...delegated, ...fresh];
    const world = buildChaosWorld(mockedAxios, { seed: SEED, users: all });

    expect(world.graph.subscriptionIdsForResourcePrefix('/me/events')).toHaveLength(40);

    // Disrupt the guard's deletes: 5 are throttled twice (must retry through), 3 hard-fail
    // (best-effort — the guard must proceed to create anyway).
    const subIdOf = (ext: string): string => world.helpers.activeDbSubsOf(ext)[0].subscriptionId;
    const flaky = delegated.slice(0, 5).map((s) => subIdOf(s.externalUserId));
    const sticky = delegated.slice(5, 8).map((s) => subIdOf(s.externalUserId));
    for (const id of flaky) world.engine.failTimes('subs.delete', id, 2, 429);
    for (const id of sticky) world.engine.alwaysFail('subs.delete', id, 500);

    const result = await drain(world.services.provisioningService.connectUsers(TENANT, inputsOf(all)));

    // Best-effort guard: even the 3 failed deletes never blocked the connect.
    expect(result).toMatchObject({ connected: 150, skipped: 0, failed: 0 });
    // Delete attempts: 32 clean (1 each) + 5 flaky (3 each) + 3 sticky (retries exhausted, 8 each).
    expect(world.metrics.attemptsFor('subs.delete')).toBe(32 + 5 * 3 + 3 * 8);
    // The 3 sticky old subs remain at Graph (delete truly failed); everything else is gone.
    expect(new Set(world.graph.subscriptionIdsForResourcePrefix('/me/events'))).toEqual(new Set(sticky));

    for (const s of delegated) {
      const subs = world.helpers.activeDbSubsOf(s.externalUserId);
      expect(subs).toHaveLength(1); // old row deactivated locally in ALL cases
      expect(subs[0].resource.startsWith('/users/')).toBe(true); // converged to app-only
    }
  });

  it('idempotency: already-connected users are skipped without any Graph traffic; a rerun converges', async () => {
    const connected: SeedUser[] = Array.from({ length: 60 }, (_, i) => ({
      externalUserId: `done-${i}`,
      email: `done-${i}@contoso.com`,
      kind: 'app-only',
      sub: { mode: 'app-only' },
    }));
    const fresh = freshUsers(90, 'todo');
    const all = [...connected, ...fresh];
    const world = buildChaosWorld(mockedAxios, {
      seed: SEED,
      users: all,
      graphRates: { throttle429: 0.1 },
      graphLatencyMs: { min: 1, max: 5 },
    });

    const run1 = await drain(world.services.provisioningService.connectUsers(TENANT, inputsOf(all)));

    expect(run1.skipped).toBe(60); // deterministic even under chaos — decided from DB pre-filter
    expect(run1.connected + run1.failed).toBe(90);
    // Skipped users generated ZERO Graph traffic.
    for (const s of connected) {
      expect(world.metrics.attemptsForKey('users.lookup', s.email)).toBe(0);
      expect(world.metrics.attemptsForKey('subs.create', world.helpers.msIdOf(s.externalUserId))).toBe(0);
    }

    // Rerun with calm weather: everything converges to connected, nothing is torn down.
    world.engine.setRates({});
    const run2 = await drain(world.services.provisioningService.connectUsers(TENANT, inputsOf(all)));
    expect(run2.skipped).toBe(60 + run1.connected);
    expect(run2.failed).toBe(0);
    expect(run2.connected).toBe(run1.failed);
    for (const s of all) {
      expect(world.helpers.activeDbSubsOf(s.externalUserId).length).toBeLessThanOrEqual(1);
    }
    // "Nothing is torn down" — no dedupe delete ever fired (skipped users are pre-filtered,
    // fresh users have nothing to remove), and the 60 seeded users still hold their ORIGINAL
    // subscriptions untouched.
    expect(world.metrics.attemptsFor('subs.delete')).toBe(0);
    for (const s of connected) {
      const subs = world.helpers.activeDbSubsOf(s.externalUserId);
      expect(subs).toHaveLength(1);
      expect(subs[0].subscriptionId.startsWith('seed-sub-')).toBe(true);
    }
  });

  it('database chaos: latency plus planned save failures — failures recorded, batch completes', async () => {
    const seeds = freshUsers(150, 'dbchaos');
    const world = buildChaosWorld(mockedAxios, {
      seed: SEED,
      users: seeds,
      dbLatencyMs: { min: 1, max: 5 },
    });

    const dbFails = seeds.slice(0, 8);
    for (const s of dbFails) world.engine.alwaysFail('db.users.save', s.externalUserId, 500);

    const result = await drain(world.services.provisioningService.connectUsers(TENANT, inputsOf(seeds)));

    expect(result.failed).toBe(8);
    expect(result.connected).toBe(142);
    for (const s of dbFails) {
      const r = result.results.find((x) => x.externalUserId === s.externalUserId);
      expect(r?.success).toBe(false);
      expect(r?.error).toContain('chaos db failure');
      // The mapping save fails BEFORE any Graph create — no external side effects for them.
      expect(world.metrics.attemptsForKey('subs.create', world.helpers.msIdOf(s.externalUserId))).toBe(0);
      expect(world.helpers.activeDbSubsOf(s.externalUserId)).toHaveLength(0);
    }
    expect(world.graph.subscriptionIdsForResourcePrefix('/users/')).toHaveLength(142);
  });

  it('host contract: POST users/connect returns before the background flow completes', async () => {
    const seeds = freshUsers(80, 'http');
    const world = buildChaosWorld(mockedAxios, {
      seed: SEED,
      users: seeds,
      graphLatencyMs: { min: 5, max: 30 },
    });

    const response = world.controllers.tenantAuthController.connectUsers({
      tenantId: TENANT,
      users: inputsOf(seeds),
    });

    // The endpoint answered synchronously; the chaotic background work has not finished.
    expect(response.totalRequested).toBe(80);
    expect(response.message).toContain('background');
    expect(world.helpers.eventsNamed(OutlookEventTypes.TENANT_USERS_BULK_CONNECT_COMPLETED)).toHaveLength(0);

    await drainUntil(
      () => world.helpers.eventsNamed(OutlookEventTypes.TENANT_USERS_BULK_CONNECT_COMPLETED).length === 1,
    );
    const summary = world.helpers.eventsNamed(OutlookEventTypes.TENANT_USERS_BULK_CONNECT_COMPLETED)[0]
      .payload as BulkConnectResult;
    // Latency-only chaos ⇒ deterministic outcome: every user actually connected.
    expect(summary).toMatchObject({ total: 80, connected: 80, failed: 0, skipped: 0 });
  });
});
