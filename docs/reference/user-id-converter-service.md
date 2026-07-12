---
dep:
  type: reference
  audience: [app-developer, ai-agent, library-contributor]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/services/shared/user-id-converter.service.ts
  tags: [identity, service, api, mapping]
  links:
    - target: ../explanation/architecture-overview.md
      rel: NEXT
---

# UserIdConverterService Reference

Injectable service that maps between the host application's `externalUserId` and the module's
internal numeric user ID. Exported from `@checkfirst/nestjs-outlook`.

## Methods

| Method | Signature | Returns |
|--------|-----------|---------|
| `toInternalUserId` | `(userId: string \| number, opts?: { cache?: boolean })` | `Promise<number>` — accepts either ID form and resolves the internal ID. |
| `externalToInternal` | `(externalUserId: string, opts?: { cache?: boolean })` | `Promise<number>` |
| `internalToExternal` | `(internalUserId: number, opts?: { cache?: boolean })` | `Promise<string>` |

## Parameters

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `cache` | `boolean` | `true` | Use the in-memory mapping cache. Pass `false` to force a fresh DB lookup. |

## Notes

- Most module operations accept an `externalUserId` and convert internally; you rarely call
  this directly unless you need the internal ID for a custom query.

## Related

- [Architecture overview](../explanation/architecture-overview.md) — why two ID spaces exist.
