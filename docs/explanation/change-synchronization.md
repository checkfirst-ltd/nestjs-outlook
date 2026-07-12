---
dep:
  type: explanation
  audience: [library-contributor, ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: medium
  depends_on:
    - src/services/shared/delta-sync.service.ts
    - src/services/calendar/recurrence.service.ts
    - src/controllers/calendar.controller.ts
    - src/enums/event-types.enum.ts
  tags: [delta, sync, webhooks, recurrence, events]
  links:
    - target: ../reference/delta-sync-service.md
      rel: EXPLAINS
    - target: ../reference/recurrence-service.md
      rel: EXPLAINS
    - target: ../reference/event-types.md
      rel: EXPLAINS
---

# Change Synchronization

This document explains how a change in a user's Outlook becomes an event your application
receives. It is conceptual background for contributors.

## From notification to event

A webhook notification from Microsoft Graph is deliberately thin: it says *something changed
for this subscription*, not *what* changed. Treating the notification as a trigger rather than
as data is the core design choice. On receipt, the module validates the notification, resolves
the affected user, and then asks Graph what actually changed. This keeps the system correct
even when notifications are coalesced, delayed, or delivered out of order — the source of truth
is always re-read, never inferred from the notification body.

## Why delta queries

Re-reading everything on every notification would be wasteful and would not scale with calendar
size. Instead the module uses Graph's delta queries: after an initial sync it stores a delta
link per user and per resource, and each subsequent fetch returns only what changed since that
link. The trade-off is extra persistent state (the delta links) and the need to handle link
invalidation — Graph can expire a delta link, at which point the module resets and re-reads from
scratch. In exchange, steady-state work per notification is bounded to the actual delta rather
than the whole mailbox. A forced reset is also used deliberately on reconnection, where the
local view may have drifted from Graph.

## Ordering changes

Changes are sorted before they are applied. Creation, update, and deletion for the same item
can arrive together, and applying them in the wrong order would leave the local view wrong.
Sorting the fetched delta into a consistent order is what lets the downstream handlers emit
coherent `EVENT_CREATED` / `EVENT_UPDATED` / `EVENT_DELETED` events. This is also why the email
deletion case produces a deletion notification followed by a creation one — the module reports
what Graph reports, in order, rather than second-guessing it.

## Recurring events

Recurring series add a dimension the raw Graph model does not hand you cleanly. A series master
carries a recurrence rule but not its concrete instances; occurrences and exceptions reference
their master. The recurrence handling normalizes these into a single shape and, when needed,
expands a series into concrete instances over a bounded window. Expansion is windowed on
purpose: an unbounded "no end" series cannot be materialized fully, so the module computes a
finite range around the present and detects which previously-stored occurrences are now stale
and should be removed. The result is a clean set — master plus instances plus a list of stale
IDs — that a consumer can persist without understanding Outlook's recurrence semantics.

## Related

- [DeltaSyncService reference](../reference/delta-sync-service.md) — the change-fetching API.
- [RecurrenceService reference](../reference/recurrence-service.md) — classification and expansion.
- [Event types reference](../reference/event-types.md) — the events this pipeline emits.
