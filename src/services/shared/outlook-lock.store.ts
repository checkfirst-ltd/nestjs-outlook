import { Logger } from "@nestjs/common";
import { randomBytes } from "crypto";
import { RedisLike } from "../../interfaces/state/redis-like.interface";

export interface WithLockResult<T> {
  acquired: boolean;
  value?: T;
}

/**
 * Per-key distributed lock with fencing tokens and TTL.
 *
 * Two backends: in-memory (single process) and Redis (cross-container).
 * Both have identical semantics — token returned on acquire must be passed
 * to renew/release, and only the holder of a matching token can renew/release.
 */
export interface OutlookLockStore {
  readonly kind: "memory" | "redis";

  /**
   * Try to acquire a lock. Returns a fencing token on success, null if held.
   * Lock auto-releases after ttlMs unless renewed.
   */
  acquireLock(key: string, ttlMs: number): Promise<string | null>;

  /**
   * Extend the lock's TTL. Returns true only if the lock is still held
   * by the caller's token. Used for heartbeat renewal.
   */
  renewLock(key: string, token: string, ttlMs: number): Promise<boolean>;

  /**
   * Release the lock iff the caller holds the matching token.
   * No-op when the token doesn't match (someone else's lock).
   */
  releaseLock(key: string, token: string): Promise<void>;
}

export function generateLockToken(): string {
  return randomBytes(16).toString("hex");
}

export class InMemoryOutlookLockStore implements OutlookLockStore {
  readonly kind = "memory" as const;
  private readonly logger = new Logger(InMemoryOutlookLockStore.name);
  private readonly locks = new Map<
    string,
    { token: string; expiresAt: number }
  >();

  async acquireLock(key: string, ttlMs: number): Promise<string | null> {
    const now = Date.now();
    const existing = this.locks.get(key);
    if (existing && existing.expiresAt > now) {
      return null;
    }
    const token = generateLockToken();
    this.locks.set(key, { token, expiresAt: now + ttlMs });
    return token;
  }

  async renewLock(key: string, token: string, ttlMs: number): Promise<boolean> {
    const now = Date.now();
    const existing = this.locks.get(key);
    if (!existing || existing.token !== token || existing.expiresAt <= now) {
      return false;
    }
    existing.expiresAt = now + ttlMs;
    return true;
  }

  async releaseLock(key: string, token: string): Promise<void> {
    const existing = this.locks.get(key);
    if (existing && existing.token === token) {
      this.locks.delete(key);
    }
  }
}

const RELEASE_LUA = `
if redis.call('GET', KEYS[1]) == ARGV[1] then
  return redis.call('DEL', KEYS[1])
else
  return 0
end
`;

const RENEW_LUA = `
if redis.call('GET', KEYS[1]) == ARGV[1] then
  return redis.call('PEXPIRE', KEYS[1], ARGV[2])
else
  return 0
end
`;

export class RedisOutlookLockStore implements OutlookLockStore {
  readonly kind = "redis" as const;
  private readonly logger = new Logger(RedisOutlookLockStore.name);

  constructor(
    private readonly redis: RedisLike,
    private readonly keyPrefix: string,
  ) {}

  private k(key: string): string {
    return `${this.keyPrefix}lock:${key}`;
  }

  async acquireLock(key: string, ttlMs: number): Promise<string | null> {
    const token = generateLockToken();
    try {
      const result = await this.redis.set(
        this.k(key),
        token,
        "PX",
        ttlMs,
        "NX",
      );
      return result === "OK" ? token : null;
    } catch (err) {
      this.logger.error(
        `[acquireLock] Redis error for ${key}: ${err instanceof Error ? err.message : String(err)}`,
      );
      return null;
    }
  }

  async renewLock(key: string, token: string, ttlMs: number): Promise<boolean> {
    try {
      const result = await this.redis.eval(
        RENEW_LUA,
        1,
        this.k(key),
        token,
        ttlMs,
      );
      return result === 1;
    } catch (err) {
      this.logger.error(
        `[renewLock] Redis error for ${key}: ${err instanceof Error ? err.message : String(err)}`,
      );
      return false;
    }
  }

  async releaseLock(key: string, token: string): Promise<void> {
    try {
      await this.redis.eval(RELEASE_LUA, 1, this.k(key), token);
    } catch (err) {
      this.logger.error(
        `[releaseLock] Redis error for ${key}: ${err instanceof Error ? err.message : String(err)}`,
      );
    }
  }
}
