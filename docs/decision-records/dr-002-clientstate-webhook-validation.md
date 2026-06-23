---
dep:
  type: decision-record
  audience: [library-contributor, ai-agent, app-developer]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: medium
  depends_on:
    - src/guards/webhook-client-state.guard.ts
    - src/controllers/calendar.controller.ts
    - src/controllers/email.controller.ts
  tags: [decision, security, webhooks, clientstate]
  links:
    - target: ../reference/http-endpoints.md
      rel: DECIDES
    - target: ../reference/event-types.md
      rel: DECIDES
---

# DR-002: Validate clientState on Webhook Endpoints

## Context

The webhook endpoints accept unauthenticated `POST` requests from Microsoft Graph — they must
be publicly reachable, so anyone on the internet can call them. Without verification, a forged
request could inject fake notifications and drive the module to act on a victim's behalf.
Microsoft Graph supports a `clientState` secret that the subscriber sets at creation time and
that Graph echoes back on every notification. (Introduced as a breaking change in #151.)

## Decision

Generate a unique `clientState` when each subscription is created, store it with the
subscription, and validate it on every inbound notification via a guard applied to all webhook
routes. Each notification item is checked individually; items whose `clientState` is missing,
references an unknown subscription, or does not match the stored value are rejected and surfaced
as a `WEBHOOK_REJECTED` event rather than processed.

## Alternatives considered

- **No verification (trust the endpoint).** Rejected: leaves the endpoints open to forged
  notifications — an unacceptable security posture for a public webhook.
- **IP allow-listing of Microsoft ranges.** Rejected: brittle (ranges change), weaker than a
  per-subscription secret, and does not bind a request to a specific subscription.
- **Validate the whole batch atomically.** Rejected in favor of per-item validation so a single
  bad item does not discard legitimate ones, and so each rejection is independently observable.

## Consequences

- Forged notifications are rejected before any side effect occurs.
- This was a breaking change: subscriptions created before the change lack a stored
  `clientState` and are rejected until recreated.
- Rejections are observable through the `WEBHOOK_REJECTED` event for monitoring.

## Review trigger

Revisit if Microsoft changes the `clientState` mechanism, if endpoint authentication moves to a
different scheme (e.g. signed tokens), or if per-item validation proves too costly at high
notification volume.
