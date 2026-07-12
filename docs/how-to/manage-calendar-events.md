---
dep:
  type: how-to
  audience: [app-developer]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/services/calendar/calendar.service.ts
  tags: [calendar, events, graph]
  links:
    - target: ../reference/calendar-service.md
      rel: USES
    - target: ../reference/permission-scopes.md
      rel: USES
    - target: ../how-to/authenticate-a-user.md
      rel: REQUIRES
---

# Manage Calendar Events

**Goal:** Create, update, fetch, and delete calendar events for an authenticated user.

These steps assume the user has already connected their Microsoft account and that you
requested `CALENDAR_WRITE` (or at least `CALENDAR_READ` for read-only steps).

## Steps

### 1. Resolve the target calendar

Most operations take a `calendarId`. Get the user's default calendar when you do not have a
specific one.

```typescript
const calendarId = await this.calendarService.getDefaultCalendarId(externalUserId);
```

### 2. Create an event

```typescript
const { event } = await this.calendarService.createEvent(
  {
    subject: 'Team Meeting',
    start: { dateTime: '2026-06-24T10:00:00', timeZone: 'UTC' },
    end: { dateTime: '2026-06-24T11:00:00', timeZone: 'UTC' },
  },
  externalUserId,
  calendarId,
);
```

### 3. Update an event

```typescript
const { event: updated } = await this.calendarService.updateEvent(
  eventId,
  { subject: 'Team Meeting (rescheduled)' },
  externalUserId,
  calendarId,
);
```

### 4. Fetch an event

```typescript
const event = await this.calendarService.getEventById(externalUserId, eventId);
```

### 5. Delete an event

```typescript
await this.calendarService.deleteEvent({ id: eventId }, externalUserId, calendarId);
```

## Verify

- `createEvent` resolves with an `event` object whose `id` is set.
- `getEventById` returns the event you created; after `deleteEvent` it returns `null`.
- The event appears in the user's Outlook calendar.

## Related

- [CalendarService reference](../reference/calendar-service.md) — all methods, including batch operations.
