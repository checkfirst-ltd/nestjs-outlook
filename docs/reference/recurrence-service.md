---
dep:
  type: reference
  audience: [app-developer, ai-agent, library-contributor]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/services/calendar/recurrence.service.ts
    - src/interfaces/recurrence.interfaces.ts
  tags: [recurrence, calendar, service, api]
  links:
    - target: ./calendar-service.md
      rel: USES
    - target: ../explanation/change-synchronization.md
      rel: NEXT
---

# RecurrenceService Reference

Injectable service that classifies Outlook events and expands recurring series into concrete
instances. Exported from `@checkfirst/nestjs-outlook`. The `Event` type is the Microsoft Graph
`Event` model.

## Methods

| Method | Signature | Returns |
|--------|-----------|---------|
| `processEvent` | `(event: Event)` | `ProcessedOutlookEvent` — normalized event ready for consumers. |
| `expandRecurringSeries` | `(seriesMasterId: string, externalUserId: string, options?: ExpandRecurringSeriesOptions)` | `Promise<RecurringEventExpansionResult>` |
| `calculateExpansionWindow` | `(recurrenceRule?: RecurrenceRule)` | `ExpansionWindow[]` — the date ranges to expand. |
| `detectStaleOccurrences` | `(fetchedExternalIds: string[], existingExternalIds: string[])` | `string[]` — IDs no longer present. |
| `isSeriesMaster` | `(event: Event)` | `boolean` — true for `seriesMaster` or events with a `recurrence`. |

## Related types

| Type | Purpose |
|------|---------|
| `RecurrenceRule` | Mirrors Graph `PatternedRecurrence` (pattern + range). |
| `OutlookEventType` | `singleInstance` \| `seriesMaster` \| `occurrence` \| `exception`. |
| `ProcessedOutlookEvent` | Normalized event with `externalId`, `eventType`, times, and optional `recurrenceRule`. |
| `RecurringEventExpansionResult` | `{ seriesMaster, instances, expansionWindow, staleExternalIds }`. |
| `ExpansionWindow` | `{ startDate: Date; endDate: Date }`. |
| `ExpandRecurringSeriesOptions` | `{ existingExternalIds?: string[] }` for stale detection. |

## Notes

- The default expansion window spans several years around the present, bounded by the rule's
  range.
- `expandRecurringSeries` throws if the series master is not found in Outlook.

## Used by

- [CalendarService reference](calendar-service.md) — fetches the events expanded here.
- [Change synchronization](../explanation/change-synchronization.md) — where expansion fits.
