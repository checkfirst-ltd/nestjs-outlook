import IoRedisMock from "ioredis-mock";
import {
  InMemoryOutlookLockStore,
  OutlookLockStore,
  RedisOutlookLockStore,
} from "./outlook-lock.store";

/**
 * The same contract is asserted against both backends. ioredis-mock runs the
 * real SET NX PX and Lua release/renew semantics in-process.
 *
 * NOTE: ioredis-mock shares one global keyspace across all client instances,
 * so each backend factory must flush before use to isolate tests.
 */
const backends: Array<[string, () => Promise<OutlookLockStore>]> = [
  ["in-memory", async () => new InMemoryOutlookLockStore()],
  [
    "redis (ioredis-mock)",
    async () => {
      const client = new IoRedisMock();
      await client.flushall();
      return new RedisOutlookLockStore(client as never, "outlook:");
    },
  ],
];

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

describe.each(backends)("OutlookLockStore (%s)", (_name, makeStore) => {
  let store: OutlookLockStore;

  beforeEach(async () => {
    store = await makeStore();
  });

  it("acquires a free key and returns a token", async () => {
    const token = await store.acquireLock("k1", 60_000);
    expect(typeof token).toBe("string");
    expect(token).toBeTruthy();
  });

  it("returns null when the key is already held", async () => {
    const first = await store.acquireLock("k1", 60_000);
    expect(first).toBeTruthy();

    const second = await store.acquireLock("k1", 60_000);
    expect(second).toBeNull();
  });

  it("releaseLock with the matching token frees the key", async () => {
    const token = await store.acquireLock("k1", 60_000);
    expect(token).toBeTruthy();

    await store.releaseLock("k1", token as string);

    const reacquired = await store.acquireLock("k1", 60_000);
    expect(reacquired).toBeTruthy();
  });

  it("releaseLock with a wrong token is a no-op (fencing)", async () => {
    const token = await store.acquireLock("k1", 60_000);
    expect(token).toBeTruthy();

    await store.releaseLock("k1", "not-the-real-token");

    // Key must still be held → a fresh acquire fails.
    const stillHeld = await store.acquireLock("k1", 60_000);
    expect(stillHeld).toBeNull();

    // Real holder can still release it.
    await store.releaseLock("k1", token as string);
    const afterRealRelease = await store.acquireLock("k1", 60_000);
    expect(afterRealRelease).toBeTruthy();
  });

  it("renewLock with the matching token returns true", async () => {
    const token = await store.acquireLock("k1", 60_000);
    const renewed = await store.renewLock("k1", token as string, 60_000);
    expect(renewed).toBe(true);
  });

  it("renewLock with a wrong token returns false", async () => {
    await store.acquireLock("k1", 60_000);
    const renewed = await store.renewLock("k1", "wrong-token", 60_000);
    expect(renewed).toBe(false);
  });

  it("lets the key be re-acquired after the TTL expires", async () => {
    const token = await store.acquireLock("k1", 80);
    expect(token).toBeTruthy();

    // Immediately, the key is held.
    expect(await store.acquireLock("k1", 80)).toBeNull();

    await sleep(120);

    const reacquired = await store.acquireLock("k1", 80);
    expect(reacquired).toBeTruthy();
  });

  it("renewLock extends the TTL so the key stays held", async () => {
    const token = await store.acquireLock("k1", 120);
    expect(token).toBeTruthy();

    await sleep(70);
    expect(await store.renewLock("k1", token as string, 120)).toBe(true);

    await sleep(70); // 140ms total elapsed, but renewed at 70ms → still held
    expect(await store.acquireLock("k1", 120)).toBeNull();
  });

  it("clearLock deletes the key regardless of token (no fencing)", async () => {
    const token = await store.acquireLock("k1", 60_000);
    expect(token).toBeTruthy();

    // A caller without the token can still clear it.
    await store.clearLock("k1");

    const reacquired = await store.acquireLock("k1", 60_000);
    expect(reacquired).toBeTruthy();
  });

  it("clearLock on an absent key is a no-op", async () => {
    await expect(store.clearLock("missing")).resolves.toBeUndefined();
  });

  it("consumeFlag returns true and clears when the flag is set", async () => {
    // markSyncPending sets the flag via acquireLock (SET NX).
    expect(await store.acquireLock("flag", 60_000)).toBeTruthy();

    expect(await store.consumeFlag("flag")).toBe(true);

    // Cleared → the key is free again.
    expect(await store.acquireLock("flag", 60_000)).toBeTruthy();
  });

  it("consumeFlag returns false when the flag is absent", async () => {
    expect(await store.consumeFlag("flag")).toBe(false);
  });

  it("consumeFlag is idempotent: a second consume returns false", async () => {
    await store.acquireLock("flag", 60_000);
    expect(await store.consumeFlag("flag")).toBe(true);
    expect(await store.consumeFlag("flag")).toBe(false);
  });
});
