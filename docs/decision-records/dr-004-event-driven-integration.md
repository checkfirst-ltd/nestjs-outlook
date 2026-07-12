---
dep:
  type: decision-record
  audience: [library-contributor, ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: medium
  depends_on:
    - src/enums/event-types.enum.ts
    - src/microsoft-outlook.module.ts
  tags: [decision, events, integration, observability]
  links:
    - target: ../reference/event-types.md
      rel: DECIDES
    - target: ../explanation/architecture-overview.md
      rel: EXPLAINS
---

# DR-004: Event-Driven Integration with the Host

## Context

The module needs to tell the host when things happen — a user authenticated, a calendar event
changed, an email arrived, a refresh token was revoked, a subscription failed. It also needs to
expose operational signals (cron completions, webhook rejections) for observability (#117,
#134). The question is how the module hands these off without coupling itself to the host's code.

## Decision

Emit typed events on the NestJS event emitter (`OutlookEventTypes`) rather than invoking
host-supplied callbacks or interfaces. The host subscribes with `@OnEvent`. Both domain changes
and lifecycle/observability signals flow through the same mechanism.

## Alternatives considered

- **Host-provided callback interfaces.** Rejected: the module would have to know the host's
  handler shapes, every new signal would change the public interface, and one slow handler could
  block the producer.
- **Return values / polling.** Rejected: webhook-driven changes are inherently asynchronous and
  push-shaped; polling adds latency and load.
- **A message broker (e.g. SQS/Kafka).** Rejected as a default: too heavy a dependency for an
  in-process library; the host can bridge events to a broker if it wants.

## Consequences

- The module and host are decoupled: new event producers can be added without changing
  consumers, and the module never needs to know who is listening.
- The link between emitter and handler is not checked at compile time — a missing or mistyped
  listener fails silently, so event names are published as an enum to reduce drift.
- Event delivery is in-process; durability across restarts is the host's concern if needed.

## Review trigger

Revisit if guaranteed delivery or cross-process fan-out becomes a requirement, or if the number
of event types grows large enough to warrant namespacing or versioning the event contract.
