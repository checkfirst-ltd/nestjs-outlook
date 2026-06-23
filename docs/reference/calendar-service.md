---
dep:
  type: reference
  audience: [app-developer, ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/services/calendar/calendar.service.ts
  tags: [calendar, service, events, api]
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

**Returns:** `Promise<{ event: Event }>`. Requests immutable IDs; retries on transient errors.

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

## Batch methods

| Method | Signature | Returns |
|--------|-----------|---------|
| `createBatchEvents` | `(events, externalUserId, calendarId)` | Per-item create results. |
| `updateBatchEvents` | `(updates, externalUserId, calendarId)` | Per-item update results. |
| `deleteBatchEvents` | `(eventIds, externalUserId, calendarId)` | Per-item delete results. |
| `getEventsBatch` | `(externalUserId, eventIds)` | Multiple events (max 20 IDs per call). |

Batch methods chunk requests into Graph `$batch` calls (20 items per batch).

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
