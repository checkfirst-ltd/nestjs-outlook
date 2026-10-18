---
dep:
  type: reference
  audience: [app-developer, ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-10-16
  last_verified: 2026-10-16T12:00:00+03:00
  confidence: high
  depends_on:
    - src/utils/all-day-event.util.ts
  tags: [calendar, all-day, graph, helpers]
  links:
    - target: ./calendar-service.md
      rel: USES
    - target: ./show-as-type.md
      rel: USES
    - target: ../how-to/manage-calendar-events.md
      rel: NEXT
---

# All-day event helpers

Utilities exported from `@checkfirst/nestjs-outlook` for building Microsoft Graph **all-day**
`start` / `end` values. They do **not** set `isAllDay` or `showAs` — those are separate fields on
the [event resource](https://learn.microsoft.com/en-us/graph/api/resources/event) (same as in the
Outlook UI).

## `buildAllDayDateTimeRange(input)`

Returns `{ start, end }` as `DateTimeTimeZone` objects at local midnight in the given
`timeZone`. Graph expects the **end** instant on the day **after** the last included calendar day
(exclusive end).

| Field | Type | Description |
|-------|------|-------------|
| `startDate` | `string` | First included day — `YYYY-MM-DD` or ISO 8601 (date portion used) |
| `endDateInclusive` | `string` | Optional last included day; defaults to `startDate` |
| `timeZone` | `string` | IANA zone for midnight boundaries (e.g. `UTC`, `Europe/London`) |

**Example — single day:**

```typescript
buildAllDayDateTimeRange({ startDate: '2026-06-25', timeZone: 'UTC' });
// start: 2026-06-25T00:00:00, end: 2026-06-26T00:00:00
```

**Example — 25–27 June inclusive:**

```typescript
buildAllDayDateTimeRange({
  startDate: '2026-06-25',
  endDateInclusive: '2026-06-27',
  timeZone: 'UTC',
});
// end: 2026-06-28T00:00:00
```

Combine with `isAllDay: true` and optional `showAs` (e.g. `ShowAsType.FREE`) on
`CalendarService.createEvent` / `updateEvent`. See [ShowAsType](show-as-type.md).

`CalendarService` also runs `normalizeAllDayEventTimes` on create/update when `isAllDay` is
true, so midnight `start` / `end` are enforced and `Z` suffixes are stripped (Graph rejects
`dateTime` values like `2026-10-21T00:00:00.000Z` when `timeZone` is set separately).

## Other exports

| Function | Purpose |
|----------|---------|
| `normalizeAllDayEventTimes` | Fix `start` / `end` on an all-day event payload before Graph |
| `calendarDateFromIso` | Parse `YYYY-MM-DD` or ISO string to a calendar date |
| `assertCalendarDate` | Validate `YYYY-MM-DD` |
| `addCalendarDays` | Calendar-date arithmetic |

## Related

- [Manage calendar events](../how-to/manage-calendar-events.md)
- [CalendarService reference](calendar-service.md)
