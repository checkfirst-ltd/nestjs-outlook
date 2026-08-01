import { Logger } from '@nestjs/common';
import axios from 'axios';
import {
  buildChaosWorld,
  SeedUser,
  AxiosMockLike,
  installChaosTimers,
  uninstallChaosTimers,
  drain,
} from '../../mock/chaos';
import { BulkConnectResult } from './tenant-provisioning.service';

jest.mock('axios');
const mockedAxios = axios as jest.Mocked<typeof axios> as unknown as AxiosMockLike;

jest.setTimeout(120_000);

/**
 * CONCURRENT-REPLICA chaos tests — the case the single-process suite structurally cannot see.
 *
 * Two independent service stacks (each = one ECS task, with its OWN in-process input dedupe)
 * run against ONE shared Microsoft Graph tenant + ONE shared RDS. This is what production looks
 * like when the service is scaled to >1 ECS task and two tasks pick up overlapping work
 * (a retried queue message, an at-least-once dispatch, or a client that double-submits).
 *
 * `connectUsers` guards duplicates with (a) per-CALL input dedupe and (b) a check-then-act
 * "already connected?" read — neither spans processes. There is no distributed lock on this
 * path (OutlookLockStore is only wired into the auth/revocation flow). So these tests DOCUMENT
 * the resulting cross-replica race, the way the rest of the suite documents production warts.
 */
describe('TenantProvisioningService — concurrent replicas (multi-process race)', () => {
  const SEED = Number(process.env.CHAOS_SEED ?? 20260714);
  const TENANT = 'tenant-t1-guid';

  const freshUsers = (n: number, prefix = 'ov'): SeedUser[] =>
    Array.from({ length: n }, (_, i) => ({
      externalUserId: `${prefix}-${i}`,
      email: `${prefix}-${i}@contoso.com`,
      kind: 'bare' as const,
      inDb: false,
    }));

  const inputsOf = (seeds: SeedUser[]) =>
    seeds.map((s) => ({ externalUserId: s.externalUserId, email: s.email }));

  const summarise = (r: BulkConnectResult | null) =>
    r ? { total: r.total, connected: r.connected, skipped: r.skipped, failed: r.failed } : null;

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

  it('two replicas connecting the SAME 40 users concurrently — measures the cross-process race', async () => {
    const N = 40;
    const seeds = freshUsers(N);
    // A small DB latency widens the check-then-act window so the interleaving is deterministic
    // under fake timers (virtual time — costs no real wall-clock).
    const replicaA = buildChaosWorld(mockedAxios, { seed: SEED, users: seeds, dbLatencyMs: { min: 2, max: 8 } });
    // Replica B: a SECOND stack over replicaA's shared Graph + DB. Its own dedupe, its own
    // "already connected" snapshot — blind to replica A, exactly like a separate ECS task.
    const replicaB = buildChaosWorld(mockedAxios, { seed: SEED + 1, reuse: replicaA.substrate });

    const [ra, rb] = await drain(
      Promise.all([
        replicaA.services.provisioningService.connectUsers(TENANT, inputsOf(seeds)),
        replicaB.services.provisioningService.connectUsers(TENANT, inputsOf(seeds)),
      ]),
    );

    // Per-user end state across the SHARED db.
    const dbSubCounts = seeds.map((s) => replicaA.helpers.activeDbSubsOf(s.externalUserId).length);
    const duplicated = dbSubCounts.filter((c) => c > 1).length;
    const orphanedZero = dbSubCounts.filter((c) => c === 0).length;
    const exactlyOne = dbSubCounts.filter((c) => c === 1).length;
    const graphSubs = replicaA.graph.subscriptionIdsForResourcePrefix('/users/').length;

    console.log(replicaA.metrics.report(`concurrent overlap N=${N} ×2 replicas seed=${SEED}`));
    console.log(`replicaA result   : ${JSON.stringify(summarise(ra))}`);
    console.log(`replicaB result   : ${JSON.stringify(summarise(rb))}`);
    console.log(`db subs/user      : exactly-1=${exactlyOne}  duplicated(>1)=${duplicated}  zero=${orphanedZero}`);
    console.log(`graph /users/ subs: ${graphSubs}  (a race-free run would be exactly ${N})`);

    // Both replicas independently believe they connected the whole batch (neither saw the other).
    expect(ra.connected).toBe(N);
    expect(rb.connected).toBe(N);

    // The invariant a correct distributed system MUST hold: exactly one live mailbox per user.
    // We assert the OBSERVED (racy) reality so the wart is codified and regressions are visible.
    // If a distributed lock is later added to this path, this expectation should flip to `toBe(0)`.
    expect(duplicated + orphanedZero).toBeGreaterThan(0); // the race is real and reproducible
    console.log(
      `\n>>> CROSS-REPLICA DEFECT: ${duplicated} user(s) hold duplicate live subscriptions, ` +
        `${orphanedZero} left with none. Graph over-provisioned by ${graphSubs - N}.`,
    );
  });

  it('two replicas on DISJOINT batches do NOT interfere (isolation sanity check)', async () => {
    const aSeeds = freshUsers(30, 'a');
    const bSeeds = freshUsers(30, 'b');
    const replicaA = buildChaosWorld(mockedAxios, {
      seed: SEED,
      users: [...aSeeds, ...bSeeds],
      dbLatencyMs: { min: 1, max: 5 },
    });
    const replicaB = buildChaosWorld(mockedAxios, { seed: SEED + 1, reuse: replicaA.substrate });

    const [ra, rb] = await drain(
      Promise.all([
        replicaA.services.provisioningService.connectUsers(TENANT, inputsOf(aSeeds)),
        replicaB.services.provisioningService.connectUsers(TENANT, inputsOf(bSeeds)),
      ]),
    );

    expect(ra.connected).toBe(30);
    expect(rb.connected).toBe(30);

    // Disjoint work → no shared user → every user holds exactly one subscription.
    for (const s of [...aSeeds, ...bSeeds]) {
      expect(replicaA.helpers.activeDbSubsOf(s.externalUserId)).toHaveLength(1);
    }
    expect(replicaA.graph.subscriptionIdsForResourcePrefix('/users/')).toHaveLength(60);
    console.log(replicaA.metrics.report(`concurrent disjoint 30+30 ×2 replicas seed=${SEED}`));
  });
});
