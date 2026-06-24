---
dep:
  type: reference
  audience: [app-developer, ai-agent, library-contributor]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/enums/microsoft-tenant-status.enum.ts
    - src/entities/microsoft-tenant.entity.ts
  tags: [enum, tenant, status, lifecycle, app-only]
  links:
    - target: ./app-only-auth-service.md
      rel: USES
    - target: ./microsoft-tenant-entity.md
      rel: USES
    - target: ../how-to/connect-enterprise-tenant.md
      rel: NEXT
    - target: ../decision-records/dr-006-dual-auth-architecture.md
      rel: NEXT
---

# MicrosoftTenantStatus Reference

`MicrosoftTenantStatus` represents the lifecycle state of a Microsoft tenant connection
for app-only authentication. Persisted on the `MicrosoftTenant` entity. Exported from
`@checkfirst/nestjs-outlook`.

## Values

| Member | String value | Meaning |
|--------|--------------|---------|
| `PENDING_CONSENT` | `PENDING_CONSENT` | Tenant registered but admin consent not yet granted. Cannot acquire tokens. |
| `ACTIVE` | `ACTIVE` | Admin consent granted, certificate valid, tokens can be acquired. |
| `CONSENT_REVOKED` | `CONSENT_REVOKED` | Admin revoked consent in Azure AD. Re-consent required. |
| `CERTIFICATE_EXPIRED` | `CERTIFICATE_EXPIRED` | Certificate has expired. Upload new certificate to Azure AD. |
| `DISABLED` | `DISABLED` | Tenant connection disabled by system or admin. No tokens issued. |

## State Transitions

```
PENDING_CONSENT ──(admin grants consent)──> ACTIVE
       │
       └──(consent denied)──> (remains PENDING_CONSENT)

ACTIVE ──(admin revokes consent)──> CONSENT_REVOKED
       │
       ├──(certificate expires)──> CERTIFICATE_EXPIRED
       │
       └──(admin disables)──> DISABLED

CONSENT_REVOKED ──(admin re-grants)──> ACTIVE

CERTIFICATE_EXPIRED ──(new cert uploaded)──> ACTIVE

DISABLED ──(admin re-enables)──> ACTIVE
```

## Notes

- Only tenants with `ACTIVE` status can acquire app-only access tokens.
- The `PENDING_CONSENT` status is the default when a tenant is first registered.
- `CONSENT_REVOKED` and `CERTIFICATE_EXPIRED` are recoverable states that require
  administrative action.
- `DISABLED` is an explicit administrative action, not an error state.

## Used by

- [AppOnlyAuthService reference](app-only-auth-service.md) — checks status before token acquisition.
- [MicrosoftTenant entity](microsoft-tenant-entity.md) — stores the status field.
- [Connect enterprise tenant](../how-to/connect-enterprise-tenant.md) — admin consent flow.
