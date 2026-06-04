import { Logger } from "@nestjs/common";
import { RedisLike } from "../../interfaces/state/redis-like.interface";

export type RateLimitWindowKey = "sec" | "min10";

export type CircuitBreakerState = "closed" | "open" | "half-open";

export interface CircuitBreakerSnapshot {
  state: CircuitBreakerState;
  openedAt: number | null;
}

/**
 * Per-user sliding-window counter + cooldown + circuit-breaker state.
 *
 * Two backends: in-memory (single process) and Redis (cross-container).
 * The Redis impl uses an atomic Lua script for the window so concurrent
 * containers can't both observe "under limit" and both push past it.
 *
 * On error, implementations return safe defaults — `Infinity` for counters
 * (caller treats as "limit hit" and waits), `null` for state reads.
 * Errors are never thrown to the caller.
 */
export interface OutlookRateLimitStore {
  readonly kind: "memory" | "redis";

  /**
   * Record a request and return the count in the trailing windowMs.
   * Atomic on the Redis backend.
   */
  recordRequest(
    userId: string,
    windowMs: number,
    key: RateLimitWindowKey,
  ): Promise<number>;

  /**
   * Read the current count in the trailing windowMs without recording.
   * Used by the rate-limiter wait loop.
   */
  getCount(
    userId: string,
    windowMs: number,
    key: RateLimitWindowKey,
  ): Promise<number>;

  getCooldown(userId: string): Promise<number | null>;

  setCooldown(userId: string, untilMs: number): Promise<void>;

  getCbState(): Promise<CircuitBreakerSnapshot | null>;

  setCbState(snapshot: CircuitBreakerSnapshot): Promise<void>;

  /**
   * Atomic SET NX for the half-open probe slot. Only one container's call
   * returns true within the probe TTL.
   */
  tryClaimHalfOpenProbe(ttlMs: number): Promise<boolean>;

  /**
   * Snapshot of active user keys (for stats). May be approximate on Redis.
   */
  getActiveUserCount(): Promise<number>;

  /**
   * Cleanup hook for in-memory backend. Redis backend uses TTLs and no-ops.
   */
  cleanupInactive(thresholdMs: number): Promise<number>;
}

interface UserState {
  windows: Map<RateLimitWindowKey, number[]>;
  cooldownUntil: number | null;
  lastActivity: number;
}

export class InMemoryOutlookRateLimitStore implements OutlookRateLimitStore {
  readonly kind = "memory" as const;
  private readonly logger = new Logger(InMemoryOutlookRateLimitStore.name);
  private readonly users = new Map<string, UserState>();
  private cbState: CircuitBreakerSnapshot = {
    state: "closed",
    openedAt: null,
  };
  private halfOpenProbeUntil = 0;

  private getOrCreate(userId: string): UserState {
    let s = this.users.get(userId);
    if (!s) {
      s = {
        windows: new Map(),
        cooldownUntil: null,
        lastActivity: Date.now(),
      };
      this.users.set(userId, s);
    }
    return s;
  }

  private trim(arr: number[], cutoff: number): number[] {
    let i = 0;
    while (i < arr.length && arr[i] < cutoff) i++;
    return i === 0 ? arr : arr.slice(i);
  }

  // Bodies are synchronous (in-process Maps); they return resolved Promises to
  // satisfy the async OutlookRateLimitStore contract whose Redis backend awaits.
  recordRequest(
    userId: string,
    windowMs: number,
    key: RateLimitWindowKey,
  ): Promise<number> {
    const now = Date.now();
    const s = this.getOrCreate(userId);
    s.lastActivity = now;
    const arr = s.windows.get(key) ?? [];
    const trimmed = this.trim(arr, now - windowMs);
    trimmed.push(now);
    s.windows.set(key, trimmed);
    return Promise.resolve(trimmed.length);
  }

  getCount(
    userId: string,
    windowMs: number,
    key: RateLimitWindowKey,
  ): Promise<number> {
    const s = this.users.get(userId);
    if (!s) return Promise.resolve(0);
    const arr = s.windows.get(key);
    if (!arr) return Promise.resolve(0);
    const now = Date.now();
    const trimmed = this.trim(arr, now - windowMs);
    s.windows.set(key, trimmed);
    return Promise.resolve(trimmed.length);
  }

  getCooldown(userId: string): Promise<number | null> {
    const s = this.users.get(userId);
    if (!s || s.cooldownUntil === null) return Promise.resolve(null);
    if (s.cooldownUntil <= Date.now()) {
      s.cooldownUntil = null;
      return Promise.resolve(null);
    }
    return Promise.resolve(s.cooldownUntil);
  }

  setCooldown(userId: string, untilMs: number): Promise<void> {
    const s = this.getOrCreate(userId);
    s.cooldownUntil = untilMs;
    s.lastActivity = Date.now();
    return Promise.resolve();
  }

  getCbState(): Promise<CircuitBreakerSnapshot | null> {
    return Promise.resolve({ ...this.cbState });
  }

  setCbState(snapshot: CircuitBreakerSnapshot): Promise<void> {
    this.cbState = { ...snapshot };
    return Promise.resolve();
  }

  tryClaimHalfOpenProbe(ttlMs: number): Promise<boolean> {
    const now = Date.now();
    if (this.halfOpenProbeUntil > now) return Promise.resolve(false);
    this.halfOpenProbeUntil = now + ttlMs;
    return Promise.resolve(true);
  }

  getActiveUserCount(): Promise<number> {
    return Promise.resolve(this.users.size);
  }

  cleanupInactive(thresholdMs: number): Promise<number> {
    const now = Date.now();
    let removed = 0;
    for (const [userId, state] of this.users.entries()) {
      if (now - state.lastActivity > thresholdMs) {
        this.users.delete(userId);
        removed++;
      }
    }
    return Promise.resolve(removed);
  }
}

const RECORD_LUA = `
local key = KEYS[1]
local now = tonumber(ARGV[1])
local windowMs = tonumber(ARGV[2])
local ttlMs = tonumber(ARGV[3])
local member = ARGV[4]
redis.call('ZREMRANGEBYSCORE', key, '-inf', now - windowMs)
redis.call('ZADD', key, now, member)
redis.call('PEXPIRE', key, ttlMs)
return redis.call('ZCARD', key)
`;

const COUNT_LUA = `
local key = KEYS[1]
local now = tonumber(ARGV[1])
local windowMs = tonumber(ARGV[2])
redis.call('ZREMRANGEBYSCORE', key, '-inf', now - windowMs)
return redis.call('ZCARD', key)
`;

export class RedisOutlookRateLimitStore implements OutlookRateLimitStore {
  readonly kind = "redis" as const;
  private readonly logger = new Logger(RedisOutlookRateLimitStore.name);
  private memberCounter = 0;

  constructor(
    private readonly redis: RedisLike,
    private readonly keyPrefix: string,
  ) {}

  private rlKey(userId: string, key: RateLimitWindowKey): string {
    return `${this.keyPrefix}rl:${userId}:${key}`;
  }

  private cdKey(userId: string): string {
    return `${this.keyPrefix}cd:${userId}`;
  }

  private cbKey(): string {
    return `${this.keyPrefix}cb`;
  }

  private cbProbeKey(): string {
    return `${this.keyPrefix}cb:probe`;
  }

  private uniqueMember(): string {
    this.memberCounter = (this.memberCounter + 1) % 1_000_000;
    return `${Date.now()}-${process.pid}-${this.memberCounter}`;
  }

  async recordRequest(
    userId: string,
    windowMs: number,
    key: RateLimitWindowKey,
  ): Promise<number> {
    try {
      const now = Date.now();
      const ttlMs = windowMs + 60_000;
      const result: unknown = await this.redis.eval(
        RECORD_LUA,
        1,
        this.rlKey(userId, key),
        now,
        windowMs,
        ttlMs,
        this.uniqueMember(),
      );
      return typeof result === "number" ? result : Number(result);
    } catch (err) {
      this.logger.error(
        `[recordRequest] Redis error for ${userId}/${key}: ${err instanceof Error ? err.message : String(err)}`,
      );
      return Number.POSITIVE_INFINITY;
    }
  }

  async getCount(
    userId: string,
    windowMs: number,
    key: RateLimitWindowKey,
  ): Promise<number> {
    try {
      const now = Date.now();
      const result: unknown = await this.redis.eval(
        COUNT_LUA,
        1,
        this.rlKey(userId, key),
        now,
        windowMs,
      );
      return typeof result === "number" ? result : Number(result);
    } catch (err) {
      this.logger.error(
        `[getCount] Redis error for ${userId}/${key}: ${err instanceof Error ? err.message : String(err)}`,
      );
      return Number.POSITIVE_INFINITY;
    }
  }

  async getCooldown(userId: string): Promise<number | null> {
    try {
      const v: unknown = await this.redis.get(this.cdKey(userId));
      if (!v) return null;
      const until = Number(v);
      if (!Number.isFinite(until) || until <= Date.now()) return null;
      return until;
    } catch (err) {
      this.logger.error(
        `[getCooldown] Redis error for ${userId}: ${err instanceof Error ? err.message : String(err)}`,
      );
      return null;
    }
  }

  async setCooldown(userId: string, untilMs: number): Promise<void> {
    try {
      const ttl = Math.max(1, untilMs - Date.now());
      await this.redis.set(this.cdKey(userId), String(untilMs), "PX", ttl);
    } catch (err) {
      this.logger.error(
        `[setCooldown] Redis error for ${userId}: ${err instanceof Error ? err.message : String(err)}`,
      );
    }
  }

  async getCbState(): Promise<CircuitBreakerSnapshot | null> {
    try {
      const hash = (await this.redis.hgetall(this.cbKey())) as
        | Record<string, string>
        | null
        | undefined;
      if (!hash || Object.keys(hash).length === 0) {
        return { state: "closed", openedAt: null };
      }
      const state = (hash.state as CircuitBreakerState | undefined) ?? "closed";
      const openedAtRaw = hash.openedAt;
      const openedAt =
        openedAtRaw && openedAtRaw !== "null" ? Number(openedAtRaw) : null;
      return { state, openedAt: Number.isFinite(openedAt) ? openedAt : null };
    } catch (err) {
      this.logger.error(
        `[getCbState] Redis error: ${err instanceof Error ? err.message : String(err)}`,
      );
      return null;
    }
  }

  async setCbState(snapshot: CircuitBreakerSnapshot): Promise<void> {
    try {
      await this.redis.hset(
        this.cbKey(),
        "state",
        snapshot.state,
        "openedAt",
        snapshot.openedAt === null ? "null" : String(snapshot.openedAt),
      );
    } catch (err) {
      this.logger.error(
        `[setCbState] Redis error: ${err instanceof Error ? err.message : String(err)}`,
      );
    }
  }

  async tryClaimHalfOpenProbe(ttlMs: number): Promise<boolean> {
    try {
      const result: unknown = await this.redis.set(
        this.cbProbeKey(),
        "1",
        "PX",
        ttlMs,
        "NX",
      );
      return result === "OK";
    } catch (err) {
      this.logger.error(
        `[tryClaimHalfOpenProbe] Redis error: ${err instanceof Error ? err.message : String(err)}`,
      );
      return false;
    }
  }

  getActiveUserCount(): Promise<number> {
    // Approximated via Redis is expensive; stats not load-bearing here.
    return Promise.resolve(0);
  }

  cleanupInactive(_thresholdMs: number): Promise<number> {
    // Redis TTLs handle eviction; no-op.
    return Promise.resolve(0);
  }
}
