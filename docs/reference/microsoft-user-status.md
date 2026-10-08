---
dep:
  type: reference
  audience:
    - app-developer
    - ai-agent
    - library-contributor
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-10-08T10:12:19.031+03:00
  confidence: high
  depends_on:
    - src/enums/microsoft-user-status.enum.ts
    - src/entities/microsoft-user.entity.ts
    - src/services/auth/microsoft-auth.service.ts
    - src/services/health/health.service.ts
  tags:
    - enum
    - user
    - status
    - lifecycle
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
| `CORRUPTED` | `CORRUPTED` | The user's **delegated** refresh token is dead (`invalid_grant`, `interaction_required`, …); delegated sync cannot work until they re-authenticate. Never set on tenant-mapped (app-only) users. |
| `SUBSCRIPTION_FAILED` | `SUBSCRIPTION_FAILED` | Webhook subscription setup failed (e.g. a `403` from Graph during validation). |

## Notes

- Only `ACTIVE` users are treated as available for token retrieval and subscription renewal
  unless `includeInactive` is requested.
- A user is moved to `SUBSCRIPTION_FAILED` when Graph rejects subscription validation, so the
  module can retry or surface the problem rather than silently dropping notifications.
- `CORRUPTED` describes delegated auth only. A tenant-mapped user (tenant **and**
  `microsoftUserId` set — the same test `resolveGraphAuth` uses to pick the app-only path)
  syncs on the tenant's token, so a dead leftover delegated token never flags them, and a
  stale flag from before their tenant mapping is ignored by the user health check and the
  subscription health cron. Rows with a tenant but no `microsoftUserId` still sync on the
  delegated token and are treated as delegated.

## Used by

- [MicrosoftSubscriptionService reference](subscription-service.md) — gates renewal on status.
- [Architecture overview](../explanation/architecture-overview.md).
