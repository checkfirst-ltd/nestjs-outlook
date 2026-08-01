import { Logger } from '@nestjs/common';
import axios from 'axios';
import {
  buildChaosWorld,
  ChaosWorld,
  SeedUser,
  AxiosMockLike,
  installChaosTimers,
  uninstallChaosTimers,
  drain,
} from '../../mock/chaos';

jest.mock('axios');
const mockedAxios = axios as jest.Mocked<typeof axios> as unknown as AxiosMockLike;

jest.setTimeout(120_000);

/**
 * Chaos + scale tests for the calendar SUBSCRIPTION lifecycle, run against BOTH auth models:
 *
 *  - `delegated`  → MicrosoftSubscriptionService.createWebhookSubscription (per-user OAuth,
 *                   MicrosoftAuthService token, resource `/me/events`).
 *  - `app-only`   → MicrosoftSubscriptionService.createAppOnlyWebhookSubscription (tenant-wide,
 *                   AppOnlyAuthService token, resource `/users/{id}/events`).
 *
 * Both drive the SAME real service + the SAME Graph POST /subscriptions boundary (behind the
 * chaos layer), so the resilience guarantees a scaled deployment relies on — tally conservation,
 * the retry ceiling, no local duplicates, and the documented at-least-once Graph duplicate on a
 * lost response — are asserted identically for each auth model via `describe.each`.
 *
 * The existing tenant-provisioning / disconnect / health chaos suites cover the tenant
 * orchestration; this suite is the auth-mode-symmetric view of the subscription create path.
 */

interface ModeSpec {
  label: 'delegated' | 'app-only';
  seedKind: SeedUser['kind'];
  /** Graph resource prefix a created subscription of this mode carries. */
  resourcePrefix: '/me/' | '/users/';
  /** The ChaosEngine plan key for this user's `subs.create` route. */
  createKey: (world: ChaosWorld, externalUserId: string) => string;
  /** Drive one subscription create through the real service for this mode. */
  create: (world: ChaosWorld, externalUserId: string) => Promise<unknown>;
}

const MODES: ModeSpec[] = [
  {
    label: 'delegated',
    seedKind: 'delegated',
    resourcePrefix: '/me/',
    // Delegated createWebhookSubscription binds clientState `user_{internalId}_…`, which the
    // fake Graph maps to the plan key `me:{internalId}`.
    createKey: (world, ext) => `me:${world.helpers.internalIdOf(ext)}`,
    create: (world, ext) =>
      world.services.subscriptionService.createWebhookSubscription(ext),
  },
  {
    label: 'app-only',
    seedKind: 'app-only',
    resourcePrefix: '/users/',
    // App-only subscriptions live at `/users/{msUserId}/events`; the plan key is that msUserId.
    createKey: (world, ext) => world.helpers.msIdOf(ext),
    create: (world, ext) =>
      world.services.subscriptionService.createAppOnlyWebhookSubscription({
        tenantId: world.defaultTenantId,
        microsoftUserId: world.helpers.msIdOf(ext),
        externalUserId: ext,
        internalUserId: world.helpers.internalIdOf(ext),
      }),
  },
];

describe.each(MODES)(
  'MicrosoftSubscriptionService — $label subscription lifecycle under chaos',
  (mode) => {
    const SEED = Number(process.env.CHAOS_SEED ?? 20260714);

    /** Users that already have a microsoft_users row of the right kind (create needs it). */
    const users = (n: number, prefix = 'u'): SeedUser[] =>
      Array.from({ length: n }, (_, i) => ({
        externalUserId: `${mode.label}-${prefix}-${i}`,
        email: `${mode.label}-${prefix}-${i}@contoso.com`,
        kind: mode.seedKind,
        inDb: true,
        sub: null,
      }));

    const createAll = (world: ChaosWorld, seeds: SeedUser[]) =>
      Promise.allSettled(seeds.map((s) => mode.create(world, s.externalUserId)));

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

    it('baseline (no chaos): creates one subscription per user, exactly one Graph call each', async () => {
      const seeds = users(60);
      const world = buildChaosWorld(mockedAxios, { seed: SEED, users: seeds });

      const results = await drain(createAll(world, seeds));

      expect(results.every((r) => r.status === 'fulfilled')).toBe(true);
      // Exactly one create attempt per user — no retries with no chaos.
      expect(world.metrics.attemptsFor('subs.create')).toBe(60);

      for (const s of seeds) {
        const subs = world.helpers.activeDbSubsOf(s.externalUserId);
        expect(subs).toHaveLength(1);
        expect(subs[0].resource.startsWith(mode.resourcePrefix)).toBe(true);
        expect(world.graph.subscriptions.has(subs[0].subscriptionId)).toBe(true);
      }
      expect(
        world.graph.subscriptionIdsForResourcePrefix(mode.resourcePrefix),
      ).toHaveLength(60);

      console.log(world.metrics.report(`sub-create baseline ${mode.label} N=60 seed=${SEED}`));
    });

    it('chaos storm: conserves tallies, stays under the retry ceiling, never duplicates locally', async () => {
      const N = 120;
      const seeds = users(N, 'storm');
      const world = buildChaosWorld(mockedAxios, {
        seed: SEED,
        users: seeds,
        graphRates: {
          throttle429: 0.15,
          unavailable503: 0.05,
          serverError500: 0.05,
          networkError: 0.05,
        },
        graphLatencyMs: { min: 1, max: 20 },
      });

      const results = await drain(createAll(world, seeds));

      const fulfilled = results.filter((r) => r.status === 'fulfilled').length;
      const rejected = results.filter((r) => r.status === 'rejected').length;

      // Conservation: every create is accounted for as success or failure.
      expect(fulfilled + rejected).toBe(N);
      // Retry ceiling: executeGraphApiCall caps at maxRetries=7 → ≤8 attempts per user.
      expect(world.metrics.attemptsFor('subs.create')).toBeLessThanOrEqual(8 * N);

      // No LOCAL duplicates: a successful user holds exactly one active DB subscription; a failed
      // one holds none. (Graph-side duplicates from lost responses are covered separately below.)
      for (const s of seeds) {
        expect(world.helpers.activeDbSubsOf(s.externalUserId).length).toBeLessThanOrEqual(1);
      }
      const localSubs = seeds.reduce(
        (acc, s) => acc + world.helpers.activeDbSubsOf(s.externalUserId).length,
        0,
      );
      expect(localSubs).toBe(fulfilled);

      console.log(world.metrics.report(`sub-create storm ${mode.label} N=${N} seed=${SEED}`));
    });

    it('at-least-once: a response lost AFTER the Graph create leaves a duplicate at Graph', async () => {
      const seeds = users(30, 'alo');
      const world = buildChaosWorld(mockedAxios, { seed: SEED, users: seeds });

      // Afflict the first 8 users: their first create mutates Graph, then the response is lost
      // (network). executeGraphApiCall retries → a SECOND Graph subscription is created; the retry
      // succeeds and is persisted. Result: one local row, two subscriptions at Graph — the
      // production at-least-once wart this suite documents.
      const afflicted = seeds.slice(0, 8);
      for (const s of afflicted) {
        world.engine.failTimesAfterExecute('subs.create', mode.createKey(world, s.externalUserId), 1, 'network');
      }

      const results = await drain(createAll(world, seeds));

      expect(results.every((r) => r.status === 'fulfilled')).toBe(true);
      // Each user has exactly one LOCAL subscription (the successful retry's persisted row).
      for (const s of seeds) {
        expect(world.helpers.activeDbSubsOf(s.externalUserId)).toHaveLength(1);
      }
      // Graph is over-provisioned by exactly the afflicted count (the orphaned first creates).
      expect(
        world.graph.subscriptionIdsForResourcePrefix(mode.resourcePrefix),
      ).toHaveLength(seeds.length + afflicted.length);

      console.log(world.metrics.report(`sub-create at-least-once ${mode.label} seed=${SEED}`));
    });

    it('planned hard failures: a targeted subset fails, the rest succeed, tallies conserved', async () => {
      const seeds = users(50, 'plan');
      const world = buildChaosWorld(mockedAxios, { seed: SEED, users: seeds });

      // 12 users whose create always 500s (past the retry ceiling) → they must fail; the other 38
      // must succeed. Isolating failures is the guarantee a bulk connect relies on.
      const doomed = seeds.slice(0, 12);
      for (const s of doomed) {
        world.engine.alwaysFail('subs.create', mode.createKey(world, s.externalUserId), 500);
      }

      const results = await drain(createAll(world, seeds));

      const fulfilled = results.filter((r) => r.status === 'fulfilled').length;
      const rejected = results.filter((r) => r.status === 'rejected').length;
      expect(fulfilled).toBe(38);
      expect(rejected).toBe(12);

      for (const s of doomed) {
        expect(world.helpers.activeDbSubsOf(s.externalUserId)).toHaveLength(0);
      }
      for (const s of seeds.slice(12)) {
        expect(world.helpers.activeDbSubsOf(s.externalUserId)).toHaveLength(1);
      }

      console.log(world.metrics.report(`sub-create planned-fail ${mode.label} seed=${SEED}`));
    });
  },
);
