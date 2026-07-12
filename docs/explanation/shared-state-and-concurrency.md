---
dep:
  type: explanation
  audience: [library-contributor, ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: medium
  depends_on:
    - src/services/shared/outlook-lock.store.ts
    - src/services/shared/outlook-rate-limit.store.ts
    - src/services/shared/graph-rate-limiter.service.ts
    - src/microsoft-outlook.module.ts
  tags: [concurrency, state, locking, rate-limit, circuit-breaker, redis]
  links:
    - target: ../reference/shared-state-stores.md
      rel: EXPLAINS
    - target: ../reference/graph-rate-limiter-service.md
      rel: EXPLAINS
    - target: ../decision-records/dr-001-pluggable-shared-state-backend.md
      rel: EXPLAINS
    - target: ../decision-records/dr-005-graph-throttling-circuit-breaker.md
      rel: EXPLAINS
---

# Shared State and Concurrency

This document explains why the module externalizes certain state and how its concurrency
controls behave. It is background for contributors, not a configuration procedure.

## The problem

Two kinds of state cannot live purely inside a single process when the application scales
horizontally. The first is **mutual exclusion**: only one worker should renew a subscription or
process a particular notification batch at a time, or the same work runs twice. The second is
**rate budgeting**: Microsoft Graph throttles per user, so all workers must share one view of
how many requests a user has made recently. If each container counts independently, the
combined traffic blows past Graph's limits even though no single container thinks it did.

When everything runs in one process, in-memory maps are enough. The moment a second container
appears, those maps diverge and both invariants break. This is the concurrency bug the module
is designed to avoid.

## The store abstraction

Rather than hard-wire Redis, the module defines two store interfaces — one for locking, one for
rate limiting — each with two interchangeable implementations. The in-memory backend is the
default and is correct for a single container. The Redis backend makes the same operations
consistent across containers. Callers depend only on the interface, so the rest of the code is
identical regardless of which backend is active. The cost of this indirection is an extra
abstraction layer and the need to keep both implementations behaviorally identical; the benefit
is that the dangerous path (multi-container) and the simple path (single container) share one
code path.

## Locks with fencing tokens

The lock store issues a fencing token on acquire and requires that token to renew or release.
This prevents a slow holder whose lock already expired from later releasing a lock someone else
now holds — only the current token-holder can act. The same store doubles as a one-bit flag
primitive: a key with no TTL persists until explicitly cleared, and `consumeFlag` atomically
tests-and-clears it so a concurrent setter is never lost in a read-then-delete gap. On Redis
these guarantees come from small Lua scripts that run the compare-and-act as a single atomic
step.

## Rate limiting and the circuit breaker

The rate-limit store tracks a per-user sliding window, a per-user cooldown, and a single
service-wide circuit-breaker state. Per-user windows enforce Graph's per-user budget; the
cooldown honors a `Retry-After` after a `429`. The circuit breaker is different in scope: it is
service-level, because sustained `503`s signal that Graph itself is unhealthy, not that one
user is noisy. After enough consecutive failures within a window the breaker opens, calls are
held off for a cooldown, and then a single container is allowed a half-open probe (claimed
atomically) to test recovery before traffic resumes. This avoids a thundering herd of retries
all hammering a struggling service at once.

## Failure philosophy

The rate-limit store never throws to its caller. On a backend error it returns safe defaults —
a counter of `Infinity`, which the caller reads as "limit hit" and waits. Degrading toward
caution rather than crashing keeps a transient Redis hiccup from turning into a flood of Graph
calls. The lock store similarly fails closed: an error on acquire returns "not acquired" rather
than a false success.

## Related

- [Shared-state stores reference](../reference/shared-state-stores.md) — the exact interfaces.
- [GraphRateLimiterService reference](../reference/graph-rate-limiter-service.md) — limits and breaker constants.
- [DR-001: Pluggable shared-state backend](../decision-records/dr-001-pluggable-shared-state-backend.md).
- [DR-005: Graph throttling and circuit breaker](../decision-records/dr-005-graph-throttling-circuit-breaker.md).
