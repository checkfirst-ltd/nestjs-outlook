---
dep:
  type: reference
  audience: [app-developer, ai-agent, library-contributor]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/entities/microsoft-tenant.entity.ts
    - src/enums/microsoft-tenant-status.enum.ts
    - src/repositories/microsoft-tenant.repository.ts
  tags: [entity, tenant, database, app-only, typeorm]
  links:
    - target: ./microsoft-tenant-status.md
      rel: USES
    - target: ./microsoft-tenant-user-entity.md
      rel: NEXT
    - target: ./app-only-auth-service.md
      rel: USES
    - target: ../how-to/connect-enterprise-tenant.md
      rel: NEXT
    - target: ../decision-records/dr-006-dual-auth-architecture.md
      rel: NEXT
---

# MicrosoftTenant Entity Reference

`MicrosoftTenant` is a TypeORM entity that stores Microsoft tenant connection information
for app-only (client credentials) authentication. Each record represents one connected
Azure AD tenant that the application can access without individual user consent.

Exported from `@checkfirst/nestjs-outlook`.

## Table

**Name**: `microsoft_tenants`

## Columns

| Column | Type | Nullable | Default | Description |
|--------|------|----------|---------|-------------|
| `id` | `INTEGER` | No | auto-increment | Primary key. |
| `tenant_id` | `VARCHAR(36)` | No | — | Azure AD tenant ID (directory ID). GUID format. **Unique, indexed**. |
| `client_id` | `VARCHAR(36)` | No | — | Azure AD app registration (client) ID for app-only auth. |
| `certificate_thumbprint` | `VARCHAR(64)` | No | — | SHA-256 certificate thumbprint for JWT signing. |
| `certificate_path` | `VARCHAR(255)` | Yes | `NULL` | Path to PEM certificate file (alternative to inline). |
| `certificate_key_path` | `VARCHAR(255)` | Yes | `NULL` | Path to PEM private key file. |
| `status` | `VARCHAR(32)` | No | `PENDING_CONSENT` | Current connection status. See [MicrosoftTenantStatus](microsoft-tenant-status.md). |
| `admin_consent_granted_at` | `DATETIME` | Yes | `NULL` | Timestamp when admin granted consent. |
| `is_active` | `BOOLEAN` | No | `true` | Quick enable/disable flag. |
| `created_at` | `TIMESTAMP` | No | `now()` | Record creation timestamp. |
| `updated_at` | `TIMESTAMP` | No | `now()` | Last update timestamp. |

## Indexes

| Index Name | Columns | Type |
|------------|---------|------|
| `IDX_microsoft_tenants_tenant_id` | `tenant_id` | Unique |

## Relationships

| Relation | Target Entity | Type | Description |
|----------|---------------|------|-------------|
| (inverse) | `MicrosoftTenantUser` | OneToMany | Users mapped within this tenant. |

## Repository

`MicrosoftTenantRepository` provides cached access:

```typescript
import { MicrosoftTenantRepository } from '@checkfirst/nestjs-outlook';

// Find tenant by Azure AD tenant ID
const tenant = await tenantRepository.findByTenantId('12345678-...');

// Update status after consent
await tenantRepository.updateStatus(tenantId, MicrosoftTenantStatus.ACTIVE);

// Mark consent granted
await tenantRepository.markConsentGranted(tenantId);
```

## Migration

Created by migration `1750000000000-AddMicrosoftTenantTables`.

## Notes

- The `tenant_id` field stores the Microsoft Azure AD tenant GUID, not an internal ID.
- Certificate information can be stored as file paths (`certificate_path`, `certificate_key_path`)
  or loaded from environment variables at runtime via `AppOnlyAuthConfig`.
- The repository uses a 60-second TTL cache keyed by `tenant_id`.

## Used by

- [AppOnlyAuthService](app-only-auth-service.md) — loads tenant credentials for token acquisition.
- [TenantAuthController](../how-to/connect-enterprise-tenant.md) — processes admin consent callbacks.
- [MicrosoftTenantUser entity](microsoft-tenant-user-entity.md) — references tenant via foreign key.
