import { Test } from "@nestjs/testing";
import { GraphRateLimiterService } from "./graph-rate-limiter.service";
import {
  CircuitBreakerSnapshot,
  InMemoryOutlookRateLimitStore,
  OutlookRateLimitStore,
} from "./outlook-rate-limit.store";
import { OUTLOOK_RATE_LIMIT_STORE } from "../../constants";

const ONE_SEC = 1000;
const TEN_MIN = 600_000;

async function buildService(
  store: OutlookRateLimitStore,
): Promise<GraphRateLimiterService> {
  const moduleRef = await Test.createTestingModule({
    providers: [
      GraphRateLimiterService,
      { provide: OUTLOOK_RATE_LIMIT_STORE, useValue: store },
    ],
  }).compile();
  return moduleRef.get(GraphRateLimiterService);
}

/**
 * A fully scripted store mock — every method is a jest.fn so tests can both
 * audit calls and drive return values across successive invocations.
 */
function makeMockStore(): jest.Mocked<OutlookRateLimitStore> {
  return {
    kind: "memory",
    recordRequest: jest.fn().mockResolvedValue(1),
    getCount: jest.fn().mockResolvedValue(0),
    getCooldown: jest.fn().mockResolvedValue(null),
    setCooldown: jest.fn().mockResolvedValue(undefined),
    getCbState: jest
      .fn()
      .mockResolvedValue({ state: "closed", openedAt: null }),
    setCbState: jest.fn().mockResolvedValue(undefined),
    tryClaimHalfOpenProbe: jest.fn().mockResolvedValue(true),
    getActiveUserCount: jest.fn().mockResolvedValue(0),
    cleanupInactive: jest.fn().mockResolvedValue(0),
  } as unknown as jest.Mocked<OutlookRateLimitStore>;
}

describe("GraphRateLimiterService", () => {
  describe("acquirePermit (call auditing on a real in-memory store)", () => {
    let store: InMemoryOutlookRateLimitStore;
    let recordSpy: jest.SpyInstance;
    let service: GraphRateLimiterService;

    beforeEach(async () => {
      store = new InMemoryOutlookRateLimitStore();
      recordSpy = jest.spyOn(store, "recordRequest");
      service = await buildService(store);
    });

    it("records exactly one 'sec' and one 'min10' request for a fresh user", async () => {
      await service.acquirePermit("u1");

      expect(recordSpy).toHaveBeenCalledTimes(2);
      expect(recordSpy).toHaveBeenCalledWith("u1", ONE_SEC, "sec");
      expect(recordSpy).toHaveBeenCalledWith("u1", TEN_MIN, "min10");
    });

    it("resolves immediately when under all limits", async () => {
      const start = Date.now();
      await service.acquirePermit("u1");
      expect(Date.now() - start).toBeLessThan(100);
    });
  });

  describe("throttle gating (fake timers, scripted store)", () => {
    let store: jest.Mocked<OutlookRateLimitStore>;
    let service: GraphRateLimiterService;

    beforeEach(async () => {
      jest.useFakeTimers();
      store = makeMockStore();
      service = await buildService(store);
    });

    afterEach(() => {
      jest.useRealTimers();
    });

    it("waits when the 1-second window is full, then records once cleared", async () => {
      // First poll: 1s window is full (>= 4). Second poll: clear.
      store.getCount
        .mockResolvedValueOnce(4) // sec — full
        .mockResolvedValueOnce(0) // min10 (first loop reads both)
        .mockResolvedValue(0); // subsequent polls clear

      const permit = service.acquirePermit("u1");

      // Let the wait loop run; advance past the 100ms internal delay.
      await jest.advanceTimersByTimeAsync(200);
      await permit;

      // It must have polled getCount more than once (waited at least one cycle).
      expect(store.getCount.mock.calls.length).toBeGreaterThan(2);
      // And ultimately recorded the request once cleared.
      expect(store.recordRequest).toHaveBeenCalledWith("u1", ONE_SEC, "sec");
    });

    it("does not record while a cooldown is active", async () => {
      const future = Date.now() + 50;
      store.getCooldown
        .mockResolvedValueOnce(future) // cooldown active
        .mockResolvedValue(null); // cleared after

      const permit = service.acquirePermit("u1");
      await jest.advanceTimersByTimeAsync(200);
      await permit;

      // recordRequest only fires after cooldown clears.
      expect(store.recordRequest).toHaveBeenCalledWith("u1", ONE_SEC, "sec");
      // getCooldown was polled more than once (waited through cooldown).
      expect(store.getCooldown.mock.calls.length).toBeGreaterThan(1);
    });
  });

  describe("handleRateLimitResponse", () => {
    it("sets a cooldown ~retryAfter seconds in the future", async () => {
      const store = makeMockStore();
      const service = await buildService(store);

      const before = Date.now();
      await service.handleRateLimitResponse("u1", 30);

      expect(store.setCooldown).toHaveBeenCalledTimes(1);
      const [userArg, untilArg] = store.setCooldown.mock.calls[0];
      expect(userArg).toBe("u1");
      // 30s ahead, allow a little execution slack.
      expect(untilArg).toBeGreaterThanOrEqual(before + 29_000);
      expect(untilArg).toBeLessThanOrEqual(before + 31_000);
    });
  });

  describe("circuit breaker", () => {
    it("opens after 5 consecutive 503 failures", async () => {
      const store = makeMockStore();
      store.getCbState.mockResolvedValue({ state: "closed", openedAt: null });
      const service = await buildService(store);

      for (let i = 0; i < 5; i++) {
        await service.record503Failure();
      }

      const openCall = store.setCbState.mock.calls.find(
        ([snap]: [CircuitBreakerSnapshot]) => snap.state === "open",
      );
      expect(openCall).toBeDefined();
    });

    it("does not open before the threshold", async () => {
      const store = makeMockStore();
      store.getCbState.mockResolvedValue({ state: "closed", openedAt: null });
      const service = await buildService(store);

      for (let i = 0; i < 4; i++) {
        await service.record503Failure();
      }

      const openCall = store.setCbState.mock.calls.find(
        ([snap]: [CircuitBreakerSnapshot]) => snap.state === "open",
      );
      expect(openCall).toBeUndefined();
    });

    it("closes the breaker when a half-open probe succeeds", async () => {
      const store = makeMockStore();
      store.getCbState.mockResolvedValue({
        state: "half-open",
        openedAt: null,
      });
      const service = await buildService(store);

      await service.recordSuccess();

      expect(store.setCbState).toHaveBeenCalledWith({
        state: "closed",
        openedAt: null,
      });
    });

    it("re-opens the breaker when a half-open probe fails (503)", async () => {
      const store = makeMockStore();
      store.getCbState.mockResolvedValue({
        state: "half-open",
        openedAt: null,
      });
      const service = await buildService(store);

      await service.record503Failure();

      const reopened = store.setCbState.mock.calls.find(
        ([snap]: [CircuitBreakerSnapshot]) => snap.state === "open",
      );
      expect(reopened).toBeDefined();
    });
  });
});
