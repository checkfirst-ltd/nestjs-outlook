---
dep:
  type: reference
  audience: [library-contributor, ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/services/shared/delta-sync.service.ts
    - src/enums/resource-type.enum.ts
    - src/entities/delta-link.entity.ts
  tags: [delta, sync, service, api, changes]
  links:
    - target: ./calendar-service.md
      rel: USES
    - target: ../explanation/change-synchronization.md
      rel: NEXT
---

# DeltaSyncService Reference

Injectable service that fetches incremental changes from Microsoft Graph using delta queries
and persists per-user, per-resource delta links. Exported from `@checkfirst/nestjs-outlook`.

## Methods

| Method | Signature | Returns |
|--------|-----------|---------|
| `fetchAndSortChanges` | `(client: Client, requestUrl: string, externalUserId: string, forceReset?: boolean, dateRange?: { startDate: Date; endDate: Date })` | `Promise<DeltaItem[]>` — changes since the stored delta link, ordered. |
| `initializeDeltaLink` | `(...)` | Establishes the initial delta link for a user/resource. |
| `saveDeltaLink` | `(internalUserId: number, resourceType: ResourceType, deltaLink: string)` | `Promise<void>` |
| `getDeltaLink` | `<T>(response: DeltaResponse<T>)` | `string \| null` — extracts `@odata.deltaLink` from a Graph response. |

## Notes

- `forceReset` discards the stored delta link and re-reads from scratch (used on reconnection).
- Delta links are stored per `ResourceType` (e.g. calendar) on the `OutlookDeltaLink` entity.
- A delta link can be invalidated by Graph; the service handles re-initialization.

## Related

- [CalendarService reference](calendar-service.md) — invokes change fetching on notifications.
- [Change synchronization](../explanation/change-synchronization.md) — the full pipeline.
