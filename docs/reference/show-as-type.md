---
dep:
  type: reference
  audience: [app-developer, ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/enums/show-as-type.enum.ts
  tags: [enum, calendar, free-busy]
  links:
    - target: ./calendar-service.md
      rel: USES
    - target: ../how-to/manage-calendar-events.md
      rel: NEXT
---

# ShowAsType Reference

`ShowAsType` mirrors the Microsoft Graph `FreeBusyStatus` type and represents the availability
shown for a calendar event. Exported from `@checkfirst/nestjs-outlook`.

## Values

| Member | String value | Meaning |
|--------|--------------|---------|
| `UNKNOWN` | `unknown` | Availability not known. |
| `FREE` | `free` | The user is available. |
| `TENTATIVE` | `tentative` | The user has tentatively accepted. |
| `BUSY` | `busy` | The user is busy. |
| `OOF` | `oof` | The user is out of office. |
| `WORKING_ELSEWHERE` | `workingElsewhere` | The user is working from another location. |

## Notes

- String values match Microsoft Graph exactly; assign them to an event's `showAs` field.
- See the Graph [event resource](https://learn.microsoft.com/en-us/graph/api/resources/event)
  for the canonical definition.

## Used by

- [CalendarService reference](calendar-service.md) — event payloads.
- [Manage calendar events](../how-to/manage-calendar-events.md).
