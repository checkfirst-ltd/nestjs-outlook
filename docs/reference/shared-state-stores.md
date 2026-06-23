---
dep:
  type: reference
  audience: [library-contributor, ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/services/shared/outlook-lock.store.ts
    - src/services/shared/outlook-rate-limit.store.ts
    - src/interfaces/state/redis-like.interface.ts
    - src/constants.ts
  tags: [state, redis, locking, rate-limit, stores]
  links:
    - target: ./configuration.md
      rel: USES
    - target: ../how-to/configure-redis-shared-state.md
      rel: NEXT
    - target: ../explanation/shared-state-and-concurrency.md
      rel: NEXT
    - target: ../decision-records/dr-001-pluggable-shared-state-backend.md
      rel: NEXT
---

# Shared-State Stores Reference

The module coordinates locking and rate-limit budgets through two pluggable store interfaces,
each with an in-memory and a Redis backend. Resolved via DI tokens and exported from
`@checkfirst/nestjs-outlook`.

## DI tokens

| Token | Resolves to |
|-------|-------------|
| `OUTLOOK_LOCK_STORE` | An `OutlookLockStore` (`RedisOutlookLockStore` or `InMemoryOutlookLockStore`). |
| `OUTLOOK_RATE_LIMIT_STORE` | An `OutlookRateLimitStore` (`RedisOutlookRateLimitStore` or `InMemoryOutlookRateLimitStore`). |

Backend selection is driven by [configuration](configuration.md) (`state.redis`).

## `OutlookLockStore`

Per-key distributed lock with fencing tokens and TTL. `kind: "memory" | "redis"`.

| Method | Signature | Returns |
|--------|-----------|---------|
| `acquireLock` | `(key: string, ttlMs?: number)` | `Promise<string \| null>` — fencing token, or `null` if held. Omitting `ttlMs` means no expiry (flag use). |
| `renewLock` | `(key: string, token: string, ttlMs: number)` | `Promise<boolean>` — true only if the caller still holds the lock. |
| `releaseLock` | `(key: string, token: string)` | `Promise<void>` — releases only if the token matches. |
| `clearLock` | `(key: string)` | `Promise<void>` — unconditional delete (flag use). |
| `consumeFlag` | `(key: string)` | `Promise<boolean>` — atomic test-and-clear of a one-bit flag. |

Helper: `generateLockToken(): string`. Result type: `WithLockResult<T>` = `{ acquired: boolean; value?: T }`.

## `OutlookRateLimitStore`

Per-user sliding-window counter + cooldown + circuit-breaker state. `kind: "memory" | "redis"`.
On error, implementations return safe defaults (`Infinity` for counters, `null` for state) and
never throw.

| Method | Signature | Returns |
|--------|-----------|---------|
| `recordRequest` | `(userId, windowMs, key: RateLimitWindowKey)` | `Promise<number>` — count in the trailing window (atomic on Redis). |
| `getCount` | `(userId, windowMs, key)` | `Promise<number>` — read without recording. |
| `getCooldown` | `(userId)` | `Promise<number \| null>` |
| `setCooldown` | `(userId, untilMs)` | `Promise<void>` |
| `getCbState` | `()` | `Promise<CircuitBreakerSnapshot \| null>` |
| `setCbState` | `(snapshot)` | `Promise<void>` |
| `tryClaimHalfOpenProbe` | `(ttlMs)` | `Promise<boolean>` — atomic SET NX; one container wins per probe TTL. |
| `getActiveUserCount` | `()` | `Promise<number>` (approximate on Redis). |
| `cleanupInactive` | `(thresholdMs)` | `Promise<number>` (in-memory only; Redis uses TTLs). |

Types: `RateLimitWindowKey` = `"sec" | "min10"`; `CircuitBreakerState` = `"closed" | "open" | "half-open"`;
`CircuitBreakerSnapshot` = `{ state: CircuitBreakerState; openedAt: number | null }`.

## `RedisLike`

Structural interface the host's client must satisfy. The module never imports `ioredis`.
Required methods: `ping`, `set`, `get`, `del`, `pexpire`, `eval`, `zadd`, `zremrangebyscore`,
`zcard`, `hset`, `hgetall`. Compatible clients: ioredis, ioredis-mock, redis-cluster wrappers,
or a test fake.

## Related

- [Configure Redis shared state](../how-to/configure-redis-shared-state.md).
- [Shared state and concurrency](../explanation/shared-state-and-concurrency.md) — the rationale.
- [DR-001: Pluggable shared-state backend](../decision-records/dr-001-pluggable-shared-state-backend.md).
