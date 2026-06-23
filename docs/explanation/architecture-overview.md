---
dep:
  type: explanation
  audience: [library-contributor, ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: medium
  depends_on:
    - src/microsoft-outlook.module.ts
    - src/services/auth/microsoft-auth.service.ts
    - src/services/shared/outlook-lock.store.ts
    - src/services/shared/outlook-rate-limit.store.ts
    - src/services/shared/delta-sync.service.ts
  tags: [architecture, design, internals, event-driven, state]
  links:
    - target: ../reference/configuration.md
      rel: EXPLAINS
    - target: ../reference/event-types.md
      rel: EXPLAINS
    - target: ../reference/subscription-service.md
      rel: EXPLAINS
    - target: ./shared-state-and-concurrency.md
      rel: NEXT
    - target: ./change-synchronization.md
      rel: NEXT
---

# Architecture Overview

This document explains how `@checkfirst/nestjs-outlook` is shaped internally and why. It is
written for contributors who need a mental model before changing the code, not a step-by-step
guide.

## The core idea

The module is a thin, opinionated boundary between a host NestJS application and the Microsoft
Graph API. The host owns its users and its HTTP surface; the module owns the messy parts of
Graph integration — the OAuth dance, token refresh, webhook subscriptions, change
synchronization, retries, and rate limiting — and hands the host clean events and service
methods. Everything the module exposes is designed so the host never has to talk to Graph
directly or reason about Graph's quirks.

## Layered composition

A single `MicrosoftOutlookModule` wires four functional areas plus a shared layer. The
**auth** area turns an OAuth callback into stored credentials and produces fresh access
tokens on demand. The **calendar** and **email** areas wrap Graph resource operations and
translate inbound notifications into domain events. The **subscription** area owns the
lifecycle of Graph webhook subscriptions, which are short-lived and must be renewed. The
**shared** layer holds cross-cutting concerns: identity conversion, rate limiting, locking,
and delta synchronization.

This separation exists because the four areas age and fail independently. A change to email
handling should not risk the calendar path, and the subscription lifecycle — by far the most
operationally fragile part — is isolated so its retries and cron jobs do not leak into the
request path.

## Event-driven by design

Rather than invoking host callbacks, the module emits events through `@nestjs/event-emitter`.
Inbound Graph notifications and internal lifecycle changes both become typed events that the
host subscribes to with `@OnEvent`. The tradeoff is indirection — the host must register
listeners and there is no compile-time link between emitter and handler — but the payoff is
decoupling: the host reacts to what happened without the module knowing or caring who is
listening, and new event producers can be added without changing consumers.

## Identity: external vs internal IDs

The host identifies users by its own `externalUserId`; the module stores users under an
internal numeric ID. A dedicated converter mediates between the two. This indirection keeps
the module's persistence independent of the host's user model, at the cost of an extra lookup
(which is cached) on most operations.

## Pluggable shared state

Two concerns — distributed locking and rate-limit budgeting — need to be consistent across
processes when the application runs in more than one container. The module abstracts both
behind store interfaces with two implementations: an in-memory one and a Redis-backed one.

The default is in-memory, which is correct only for a single container. When a Redis client is
supplied, the stores coordinate across containers. The `required` flag captures a deliberate
operational tradeoff: with it off, a Redis outage silently degrades to in-memory (availability
over correctness, but reintroducing the cross-container concurrency bug); with it on, a failed
Redis probe crashes module init so the orchestrator restarts the container instead of running
in a subtly wrong mode. The module never imports a Redis driver itself — the host injects a
compatible client — so the dependency stays the host's choice.

## Resilience: rate limiting, retries, and delta sync

Graph imposes throttling and returns transient failures. Calendar and email operations run
through a shared rate limiter and a retry helper, so callers get a simple promise while the
module absorbs `429`s and transient errors with backoff. Change synchronization uses Graph's
delta queries: the module persists a delta link per user and resource so that, after a
notification, it fetches only what changed rather than re-reading everything. The cost is
extra stored state (delta links) and the need to handle delta-link invalidation; the benefit
is bounded, incremental work per notification.

## Security at the edges

Two mechanisms protect the trust boundary. CSRF tokens are embedded in the OAuth `state`
parameter and validated on callback, preventing forged authorization redirects. Webhook
notifications carry a `clientState` that is generated when a subscription is created and
verified on every inbound notification; mismatches are rejected and surfaced as a security
event rather than processed. Both reflect the same principle: never trust an inbound request
that claims to be from Microsoft without proof tied to state the module itself issued.

## Related references

- [Configuration](../reference/configuration.md) — the surface that exposes these choices.
- [Event types](../reference/event-types.md) — the contract the event-driven design produces.
- [Subscription service](../reference/subscription-service.md) — the subscription lifecycle API.
- [Shared state and concurrency](shared-state-and-concurrency.md) — locking, rate limiting, and the circuit breaker.
- [Change synchronization](change-synchronization.md) — how notifications become events.
