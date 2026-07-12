---
dep:
  type: reference
  audience: [app-developer, ai-agent, library-contributor]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/enums/microsoft-user-status.enum.ts
    - src/entities/microsoft-user.entity.ts
  tags: [enum, user, status, lifecycle]
  links:
    - target: ./subscription-service.md
      rel: USES
    - target: ../explanation/architecture-overview.md
      rel: NEXT
---

# MicrosoftUserStatus Reference

`MicrosoftUserStatus` is the stored state of a connected Microsoft user, persisted on the
`MicrosoftUser` entity. Exported from `@checkfirst/nestjs-outlook`.

## Values

| Member | String value | Meaning |
|--------|--------------|---------|
| `ACTIVE` | `ACTIVE` | User is connected and usable; tokens valid and subscriptions healthy. |
| `CORRUPTED` | `CORRUPTED` | User record is in an inconsistent state and needs attention. |
| `SUBSCRIPTION_FAILED` | `SUBSCRIPTION_FAILED` | Webhook subscription setup failed (e.g. a `403` from Graph during validation). |

## Notes

- Only `ACTIVE` users are treated as available for token retrieval and subscription renewal
  unless `includeInactive` is requested.
- A user is moved to `SUBSCRIPTION_FAILED` when Graph rejects subscription validation, so the
  module can retry or surface the problem rather than silently dropping notifications.

## Used by

- [MicrosoftSubscriptionService reference](subscription-service.md) — gates renewal on status.
- [Architecture overview](../explanation/architecture-overview.md).
