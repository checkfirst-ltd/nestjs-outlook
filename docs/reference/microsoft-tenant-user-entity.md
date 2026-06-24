---
dep:
  type: reference
  audience: [app-developer, ai-agent, library-contributor]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/entities/microsoft-tenant-user.entity.ts
    - src/entities/microsoft-tenant.entity.ts
  tags: [entity, user, tenant, database, app-only, typeorm, mapping]
  links:
    - target: ./microsoft-tenant-entity.md
      rel: USES
    - target: ./tenant-user-service.md
      rel: USES
    - target: ./tenant-calendar-service.md
      rel: NEXT
    - target: ../decision-records/dr-008-tenant-user-mapping.md
      rel: NEXT
---

# MicrosoftTenantUser Entity Reference

`MicrosoftTenantUser` is a TypeORM entity that maps external user IDs (from the host
application) to Microsoft user IDs within a connected tenant. This mapping enables
app-only services to operate on specific users' calendars and emails.

Exported from `@checkfirst/nestjs-outlook`.

## Table

**Name**: `microsoft_tenant_users`

## Columns

| Column | Type | Nullable | Default | Description |
|--------|------|----------|---------|-------------|
| `id` | `INTEGER` | No | auto-increment | Primary key. |
| `tenant_id` | `INTEGER` | No | — | Foreign key to `microsoft_tenants.id`. |
| `microsoft_user_id` | `VARCHAR(36)` | No | — | Azure AD user object ID. Used for `/users/{id}/*` Graph API calls. |
| `external_user_id` | `VARCHAR(255)` | No | — | Host application's user ID. **Indexed**. |
| `user_principal_name` | `VARCHAR(255)` | No | — | UPN in format `user@tenant.onmicrosoft.com`. |
| `default_calendar_id` | `VARCHAR(255)` | Yes | `NULL` | Cached default calendar ID for performance. |
| `is_active` | `BOOLEAN` | No | `true` | Whether this mapping is active. |
| `created_at` | `TIMESTAMP` | No | `now()` | Record creation timestamp. |
| `updated_at` | `TIMESTAMP` | No | `now()` | Last update timestamp. |

## Indexes

| Index Name | Columns | Type |
|------------|---------|------|
| `IDX_microsoft_tenant_users_external_user_id` | `external_user_id` | Non-unique |
| `IDX_microsoft_tenant_users_microsoft_user_id` | `microsoft_user_id` | Non-unique |

## Relationships

| Relation | Target Entity | Type | Cascade | Description |
|----------|---------------|------|---------|-------------|
| `tenant` | `MicrosoftTenant` | ManyToOne | CASCADE (on delete) | The tenant this user belongs to. |

## TypeORM Definition

```typescript
@Entity('microsoft_tenant_users')
export class MicrosoftTenantUser {
  @PrimaryGeneratedColumn('increment')
  id!: number;

  @ManyToOne(() => MicrosoftTenant)
  @JoinColumn({ name: 'tenant_id' })
  tenant!: MicrosoftTenant;

  @Column({ name: 'microsoft_user_id', length: 36 })
  microsoftUserId: string = '';

  @Column({ name: 'external_user_id', length: 255 })
  @Index()
  externalUserId: string = '';

  @Column({ name: 'user_principal_name', length: 255 })
  userPrincipalName: string = '';

  @Column({ name: 'default_calendar_id', length: 255, nullable: true })
  defaultCalendarId: string | null = null;

  @Column({ name: 'is_active', default: true })
  isActive: boolean = true;

  // ... timestamps
}
```

## Usage

The mapping is created when a host application registers a user for tenant-wide access:

```typescript
import { TenantUserService } from '@checkfirst/nestjs-outlook';

// Register mapping from external user to Microsoft user
const tenantUser = await tenantUserService.registerUserMapping(
  'tenant-guid',           // Microsoft tenant ID
  'external-user-123',     // Your app's user ID
  'john@contoso.com'       // User's email (used to lookup Microsoft ID)
);

// Later, retrieve Microsoft user ID for Graph API calls
const msUserId = await tenantUserService.getMicrosoftUserId(
  'tenant-guid',
  'external-user-123'
);
```

## Migration

Created by migration `1750000000000-AddMicrosoftTenantTables`.

## Notes

- The `microsoft_user_id` is the Azure AD user object ID (immutable ID when `Prefer: IdType="ImmutableId"`
  header is used).
- `external_user_id` is provided by the host application and can be any string identifier.
- The `default_calendar_id` is cached after first lookup to avoid repeated Graph API calls.
- Deleting a `MicrosoftTenant` cascades to delete all associated `MicrosoftTenantUser` records.

## Used by

- [TenantUserService](tenant-user-service.md) — manages user lookup and mapping.
- [TenantCalendarService](tenant-calendar-service.md) — uses mapping for calendar operations.
- [DR-008: Tenant User Mapping](../decision-records/dr-008-tenant-user-mapping.md) — design rationale.
