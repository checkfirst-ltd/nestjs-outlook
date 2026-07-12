---
dep:
  type: reference
  audience: [library-contributor, ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/services/shared/graph-rate-limiter.service.ts
    - src/services/shared/outlook-rate-limit.store.ts
  tags: [rate-limit, throttling, circuit-breaker, service, api]
  links:
    - target: ./shared-state-stores.md
      rel: USES
    - target: ../explanation/shared-state-and-concurrency.md
      rel: NEXT
    - target: ../decision-records/dr-005-graph-throttling-circuit-breaker.md
      rel: NEXT
---

# GraphRateLimiterService Reference

Injectable service that throttles Microsoft Graph calls per user and trips a service-level
circuit breaker on sustained `503`s. Exported from `@checkfirst/nestjs-outlook`. Backed by an
`OutlookRateLimitStore`.

## Methods

| Method | Signature | Returns |
|--------|-----------|---------|
| `acquirePermit` | `(userId: string)` | `Promise<void>` — waits until the user is under the per-user limits and any cooldown has elapsed. |
| `releasePermit` | `(userId: string)` | `void` |
| `handleRateLimitResponse` | `(userId: string, retryAfterSeconds: number)` | `Promise<void>` — records a `429` and sets a cooldown. |
| `recordSuccess` | `()` | `Promise<void>` — feeds the circuit breaker's recovery. |
| `getStats` | `()` | service-level statistics. |
| `getUserStats` | `(userId: string)` | per-user statistics. |

## Limits and thresholds

| Constant | Value | Meaning |
|----------|-------|---------|
| `MAX_REQUESTS_PER_SECOND` | `4` | Per-user requests allowed per second. |
| `MAX_REQUESTS_PER_10_MINUTES` | `10000` | Per-user requests allowed per 10 minutes. |
| `INACTIVE_USER_THRESHOLD_MS` | `1800000` (30 min) | In-memory cleanup threshold for idle users. |
| `CB_FAILURE_THRESHOLD` | `5` | Consecutive `503`s that trip the breaker. |
| `CB_FAILURE_WINDOW_MS` | `60000` | Window in which those `503`s must occur. |
| `CB_COOLDOWN_MS` | `60000` | How long the breaker stays open. |
| `CB_PROBE_TTL_MS` | `5000` | Half-open probe slot TTL. |

## Notes

- Used internally by `CalendarService` and `EmailService`; you rarely call it directly.
- Window counters use the `RateLimitWindowKey` values `"sec"` and `"min10"`.

## Related

- [Shared-state stores reference](shared-state-stores.md) — the backing store interface.
- [DR-005: Graph throttling and circuit breaker](../decision-records/dr-005-graph-throttling-circuit-breaker.md).
