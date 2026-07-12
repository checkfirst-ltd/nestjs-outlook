---
dep:
  type: decision-record
  audience: [library-contributor, ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: medium
  depends_on:
    - src/microsoft-outlook.module.ts
    - src/interfaces/config/outlook-config.interface.ts
    - src/services/shared/outlook-lock.store.ts
    - src/services/shared/outlook-rate-limit.store.ts
  tags: [decision, state, redis, concurrency]
  links:
    - target: ../reference/configuration.md
      rel: DECIDES
    - target: ../reference/shared-state-stores.md
      rel: DECIDES
    - target: ../explanation/shared-state-and-concurrency.md
      rel: EXPLAINS
---

# DR-001: Pluggable Shared-State Backend (Redis or In-Memory)

## Context

The module holds two kinds of state that must be consistent across processes: distributed locks
(so work like subscription renewal runs once) and per-user rate-limit budgets (so all workers
share one view of Graph usage). In-memory state is correct for a single container but diverges
the moment a second container runs, reintroducing duplicate work and rate-limit overruns. The
module is consumed by services that deploy multiple containers, so single-process assumptions
are unsafe by default.

## Decision

Abstract locking and rate limiting behind store interfaces with two interchangeable backends —
in-memory (default) and Redis — selected through `state.redis` configuration. The host supplies
an ioredis-compatible client; the module never imports a Redis driver. A `required` flag
controls whether a failed Redis probe at startup crashes module init or falls back to in-memory.

## Alternatives considered

- **Hard dependency on ioredis.** Rejected: forces a Redis dependency on every consumer,
  including single-container and test setups that do not need it, and couples the module to one
  client library.
- **In-memory only.** Rejected: silently incorrect under horizontal scaling — the exact
  environment the module targets.
- **Always require Redis, no fallback.** Rejected: too rigid for local development and tests;
  the `required` flag lets each deployment choose its own safety/availability trade-off.

## Consequences

- The same code path serves both single- and multi-container deployments.
- Both backends must be kept behaviorally identical, which is an ongoing maintenance burden.
- With `required: false`, a Redis outage degrades to in-memory and re-opens the concurrency
  risk; this is logged but not fatal. Production deployments are expected to set `required: true`.

## Review trigger

Revisit if the module adopts a different coordination primitive (e.g. a database advisory lock),
if a non-Redis shared-state backend is needed, or if the dual-implementation burden causes
behavioral drift between backends.
