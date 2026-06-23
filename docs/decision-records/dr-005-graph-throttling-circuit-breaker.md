---
dep:
  type: decision-record
  audience: [library-contributor, ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: medium
  depends_on:
    - src/services/shared/graph-rate-limiter.service.ts
    - src/services/shared/outlook-rate-limit.store.ts
  tags: [decision, rate-limit, throttling, circuit-breaker, resilience]
  links:
    - target: ../reference/graph-rate-limiter-service.md
      rel: DECIDES
    - target: ../reference/shared-state-stores.md
      rel: DECIDES
    - target: ../explanation/shared-state-and-concurrency.md
      rel: EXPLAINS
---

# DR-005: Per-User Throttling with a Service-Level Circuit Breaker

## Context

Microsoft Graph enforces per-user throttling and returns `429` with `Retry-After` when a user
exceeds it, and `503` when the service itself is overloaded (#132). A naive client that retries
immediately makes both situations worse — it burns the user's budget faster and piles load onto
an already-struggling service.

## Decision

Gate every Graph call through a rate limiter that maintains per-user sliding-window counters and
honors `429` cooldowns, plus a single service-level circuit breaker for sustained `503`s. The
breaker opens after a threshold of consecutive failures within a window, holds calls off for a
cooldown, then allows one atomically-claimed half-open probe to test recovery before resuming.
Counters, cooldowns, and breaker state live in the shared-state store so they are consistent
across containers.

## Alternatives considered

- **Naive retry with backoff only.** Rejected: backoff alone does not coordinate across users
  or containers and keeps hammering a `503`-ing service.
- **Per-user circuit breakers.** Rejected: a `503` reflects service health, not one user's
  behavior; a service-level breaker is the correct scope. Per-user concerns are handled by the
  per-user windows and cooldowns instead.
- **Client-side throttling without shared state.** Rejected: independent per-container counters
  collectively overshoot Graph's per-user limit (see [DR-001](dr-001-pluggable-shared-state-backend.md)).

## Consequences

- Transient `429`/`503` conditions are absorbed by the module; callers see a simple promise that
  resolves once it is safe to proceed.
- During an open breaker, calls wait rather than fail fast — latency rises but the dependent
  service is given room to recover.
- Correct multi-container behavior depends on the Redis-backed store; with the in-memory backend
  the protections are per-process only.

## Review trigger

Revisit if Graph's documented limits change, if waiting-on-open proves worse than failing fast
for some call sites, or if the fixed thresholds need to become configurable.
