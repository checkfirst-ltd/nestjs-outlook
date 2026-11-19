---
dep:
  type: reference
  audience:
    - app-developer
    - ai-agent
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-11-18T22:38:11.432+02:00
  confidence: high
  depends_on:
    - src/services/calendar/calendar.service.ts
    - src/utils/all-day-event.util.ts
    - src/utils/retry.util.ts
    - src/errors/duplicate-transaction-id.error.ts
  tags:
    - calendar
    - service
    - events
    - api
  links:
    - target: ../how-to/manage-calendar-events.md
      rel: NEXT
---

# CalendarService Reference

Injectable service for calendar operations against Microsoft Graph. Exported from
`@checkfirst/nestjs-outlook`. The `Event` type is the Microsoft Graph `Event` model.

## Methods

### `getDefaultCalendarId(externalUserId)`

| Parameter | Type | Description |
|-----------|------|-------------|
| `externalUserId` | `string` | Host application user ID. |

**Returns:** `Promise<string>` — the user's default calendar ID (cached after first lookup).

### `createEvent(event, externalUserId, calendarId)`

| Parameter | Type | Description |
|-----------|------|-------------|
| `event` | `Partial<Event>` | Event payload (subject, start, end, …). |
| `externalUserId` | `string` | Host application user ID. |
| `calendarId` | `string` | Target calendar ID. |

**Returns:** `Promise<{ event: Event; adoptedExisting?: boolean }>`. Requests immutable IDs;
retries on transient errors.

A create is idempotent on `event.transactionId`. If Graph answers `ErrorDuplicateTransactionId`
(an earlier attempt already made the event, e.g. its answer was lost to a 5xx), the method
does not retry. It updates the event Graph names in `@event.existingEventId` with this payload
and returns it with `adoptedExisting: true`. If that event is gone, or Graph names none, it
throws `DuplicateTransactionIdError`.

For **all-day** events, use `buildAllDayDateTimeRange` for `start` / `end` only; set `isAllDay`
and `showAs` on the event body. See [All-day event helpers](all-day-event-helpers.md) and
[Manage calendar events](../how-to/manage-calendar-events.md).

### `updateEvent(eventId, updates, externalUserId, calendarId)`

| Parameter | Type | Description |
|-----------|------|-------------|
| `eventId` | `string` | ID of the event to update. |
| `updates` | `Partial<Event>` | Fields to change. |
| `externalUserId` | `string` | Host application user ID. |
| `calendarId` | `string` | Calendar containing the event. |

**Returns:** `Promise<{ event: Event }>`.

### `getEventById(externalUserId, eventId)`

**Returns:** `Promise<Event | null>` — the event, or `null` if not found.

### `deleteEvent(event, externalUserId, calendarId)`

| Parameter | Type | Description |
|-----------|------|-------------|
| `event` | `Partial<Event>` | Event to delete (must include `id`). |
| `externalUserId` | `string` | Host application user ID. |
| `calendarId` | `string` | Calendar containing the event. |

**Returns:** `Promise<void>`.

### `findEventsByTransactionId(externalUserId, transactionId, window)`

| Parameter | Type | Description |
|-----------|------|-------------|
| `externalUserId` | `string` | Host application user ID. |
| `transactionId` | `string` | The full transactionId the event was created with. |
| `window` | `{ start: string \| Date; end: string \| Date }` | Roughly where the event is. Times without an offset are read as UTC; the window is widened by a day each side. |

**Returns:** `Promise<Event[]>`, every event in the window with that transactionId.

Graph cannot `$filter` events by `transactionId`, so this reads the calendar view of the window
and compares. Use it to find an event whose id was never stored.

## Batch methods

| Method | Signature | Returns |
|--------|-----------|---------|
| `createBatchEvents` | `(events, externalUserId, calendarId)` | Per-item create results. |
| `updateBatchEvents` | `(updates, externalUserId, calendarId)` | Per-item update results. |
| `deleteBatchEvents` | `(eventIds, externalUserId, calendarId)` | Per-item delete results. |
| `getEventsBatch` | `(externalUserId, eventIds)` | Multiple events (max 20 IDs per call). |

Batch methods chunk requests into Graph `$batch` calls (20 items per batch). Graph may answer
the items of a batch in any order, so each answer is matched to its request by its `id`.
`createBatchEvents` adopts an item whose transactionId is taken, as `createEvent` does.

## Other methods

| Method | Signature | Returns |
|--------|-----------|---------|
| `getEventDetails` | `(...)` | Detailed event payload. |
| `getAuthenticatedClient` | `(externalUserId)` | `Promise<Client>` — a Graph client for the user. |
| `saveDeltaLink` | `(externalUserId, deltaLink)` | `Promise<void>` — persists a delta link. |
| `handleOutlookWebhook` | `(...)` | Processes an inbound calendar notification. |
| `fetchAndSortChanges` | `(...)` | Fetches and orders changes via delta sync. |

## Used by

- [Manage calendar events](../how-to/manage-calendar-events.md).
