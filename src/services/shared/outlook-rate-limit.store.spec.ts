import IoRedisMock from "ioredis-mock";
import {
  InMemoryOutlookRateLimitStore,
  OutlookRateLimitStore,
  RedisOutlookRateLimitStore,
} from "./outlook-rate-limit.store";

const backends: Array<[string, () => Promise<OutlookRateLimitStore>]> = [
  ["in-memory", async () => new InMemoryOutlookRateLimitStore()],
  [
    "redis (ioredis-mock)",
    async () => {
      const client = new IoRedisMock();
      await client.flushall();
      // Unique prefix per suite: ioredis-mock shares one global keyspace across
      // all client instances and parallel jest workers, so a shared prefix would
      // let suites clobber each other's keys. Namespacing keeps them isolated.
      return new RedisOutlookRateLimitStore(client as never, "outlook-rl-test:");
    },
  ],
];

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

describe.each(backends)("OutlookRateLimitStore (%s)", (_name, makeStore) => {
  let store: OutlookRateLimitStore;

  beforeEach(async () => {
    store = await makeStore();
  });

  describe("recordRequest / getCount", () => {
    it("returns the post-insert count, incrementing 1..N", async () => {
      expect(await store.recordRequest("u1", 1000, "sec")).toBe(1);
      expect(await store.recordRequest("u1", 1000, "sec")).toBe(2);
      expect(await store.recordRequest("u1", 1000, "sec")).toBe(3);
      expect(await store.getCount("u1", 1000, "sec")).toBe(3);
    });

    it("tracks the 'sec' and 'min10' windows independently", async () => {
      await store.recordRequest("u1", 1000, "sec");
      await store.recordRequest("u1", 600_000, "min10");
      await store.recordRequest("u1", 600_000, "min10");
      expect(await store.getCount("u1", 1000, "sec")).toBe(1);
      expect(await store.getCount("u1", 600_000, "min10")).toBe(2);
    });

    it("drops records older than the window (sliding window)", async () => {
      // Use a tiny window so records age out within the test.
      await store.recordRequest("u1", 60, "sec");
      await store.recordRequest("u1", 60, "sec");
      expect(await store.getCount("u1", 60, "sec")).toBe(2);

      await sleep(90);

      // All prior records are now older than the 60ms window.
      expect(await store.getCount("u1", 60, "sec")).toBe(0);
    });

    it("isolates counts per user", async () => {
      await store.recordRequest("u1", 1000, "sec");
      await store.recordRequest("u2", 1000, "sec");
      await store.recordRequest("u2", 1000, "sec");
      expect(await store.getCount("u1", 1000, "sec")).toBe(1);
      expect(await store.getCount("u2", 1000, "sec")).toBe(2);
    });

    it("counts exactly N under 100 concurrent records (atomicity)", async () => {
      const calls = Array.from({ length: 100 }, () =>
        store.recordRequest("u-atomic", 600_000, "min10"),
      );
      await Promise.all(calls);
      expect(await store.getCount("u-atomic", 600_000, "min10")).toBe(100);
    });
  });

  describe("cooldown", () => {
    it("returns a future cooldown and null once it has passed", async () => {
      const until = Date.now() + 120;
      await store.setCooldown("u1", until);

      const read = await store.getCooldown("u1");
      expect(read).not.toBeNull();
      expect(read).toBeGreaterThan(Date.now());

      await sleep(160);
      expect(await store.getCooldown("u1")).toBeNull();
    });

    it("returns null when no cooldown is set", async () => {
      expect(await store.getCooldown("never-set")).toBeNull();
    });
  });

  describe("circuit-breaker state", () => {
    it("defaults to closed when unset", async () => {
      const cb = await store.getCbState();
      expect(cb?.state).toBe("closed");
    });

    it("round-trips an open state", async () => {
      const openedAt = Date.now();
      await store.setCbState({ state: "open", openedAt });
      const cb = await store.getCbState();
      expect(cb?.state).toBe("open");
      expect(cb?.openedAt).toBe(openedAt);
    });

    it("round-trips a half-open state with null openedAt", async () => {
      await store.setCbState({ state: "half-open", openedAt: null });
      const cb = await store.getCbState();
      expect(cb?.state).toBe("half-open");
      expect(cb?.openedAt).toBeNull();
    });
  });

  describe("tryClaimHalfOpenProbe", () => {
    it("grants exactly one of N concurrent claims", async () => {
      const claims = await Promise.all(
        Array.from({ length: 10 }, () => store.tryClaimHalfOpenProbe(5000)),
      );
      const granted = claims.filter(Boolean).length;
      expect(granted).toBe(1);
    });

    it("grants again only after the probe TTL expires", async () => {
      expect(await store.tryClaimHalfOpenProbe(80)).toBe(true);
      expect(await store.tryClaimHalfOpenProbe(80)).toBe(false);

      await sleep(120);
      expect(await store.tryClaimHalfOpenProbe(80)).toBe(true);
    });
  });
});
