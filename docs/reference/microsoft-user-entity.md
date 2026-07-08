---
dep:
  type: reference
  audience: [app-developer, ai-agent, library-contributor]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-07-08T12:00:00+03:00
  confidence: high
  depends_on:
    - src/entities/microsoft-user.entity.ts
    - src/entities/microsoft-tenant.entity.ts
  tags: [entity, user, tenant, database, app-only, delegated, typeorm, mapping]
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

# MicrosoftUser Entity Reference

`MicrosoftUser` is the single TypeORM entity representing a host-application user. One row
(keyed by `external_user_id`) carries **either or both** authentication capabilities:

- **Delegated (per-user OAuth):** `access_token` / `refresh_token` / `token_expiry` /
  `scopes` are populated once the user completes the OAuth flow.
- **App-only (tenant-wide):** `tenant` / `microsoft_user_id` / `user_principal_name` are
  populated when the host maps the user into a Microsoft tenant via
  [`TenantUserService.registerUserMapping`](tenant-user-service.md). No per-user tokens are
  stored for this mode — the token is acquired at the tenant level — which is why the
  delegated token columns are nullable.

Keeping both capabilities on one row means shared user features (default calendar, active
flag, future per-user settings) are implemented once regardless of how the user
authenticated. See [DR-008](../decision-records/dr-008-tenant-user-mapping.md) for the
rationale behind consolidating storage onto this table.

Exported from `@checkfirst/nestjs-outlook`.

## Table

**Name**: `microsoft_users`

## Columns

| Column | Type | Nullable | Default | Description |
|--------|------|----------|---------|-------------|
| `id` | `INTEGER` | No | auto-increment | Primary key. |
| `external_user_id` | `VARCHAR(255)` | No | `''` | Host application's user ID. **Indexed**. The mapping key — one row per host user. |
| `access_token` | `TEXT` | Yes | `NULL` | Delegated OAuth access token. Null for app-only-only users. |
| `refresh_token` | `TEXT` | Yes | `NULL` | Delegated OAuth refresh token. Null for app-only-only users. |
| `token_expiry` | `DATETIME` | Yes | `NULL` | Delegated access-token expiry. |
| `scopes` | `TEXT` | Yes | `NULL` | Scopes granted at delegated authentication, reused on refresh. |
| `is_active` | `BOOLEAN` | No | `true` | Whether this user record is active. |
| `status` | `VARCHAR(32)` | No | `ACTIVE` | Delegated user lifecycle status (see [MicrosoftUserStatus](microsoft-user-status.md)). |
| `default_calendar_id` | `VARCHAR(255)` | Yes | `NULL` | Cached default calendar ID for performance. |
| `tenant_id` | `INTEGER` | Yes | `NULL` | Foreign key to `microsoft_tenants.id` for app-only access. Null for delegated-only users. |
| `microsoft_user_id` | `VARCHAR(36)` | Yes | `NULL` | Azure AD user object ID, used for `/users/{id}/*` Graph calls. **Indexed**. |
| `user_principal_name` | `VARCHAR(320)` | Yes | `NULL` | UPN, e.g. `user@tenant.onmicrosoft.com`. |
| `created_at` | `TIMESTAMP` | No | `now()` | Record creation timestamp. |
| `updated_at` | `TIMESTAMP` | No | `now()` | Last update timestamp. |

## Indexes

| Index Name | Columns | Type |
|------------|---------|------|
| `IDX_microsoft_users_external_user_id` | `external_user_id` | **Unique** — one row per host user |
| `IDX_microsoft_users_microsoft_user_id` | `microsoft_user_id` | Non-unique |

## Relationships

| Relation | Target Entity | Type | Cascade | Description |
|----------|---------------|------|---------|-------------|
| `tenant` | `MicrosoftTenant` | ManyToOne (nullable) | CASCADE (on delete) | The tenant this user is mapped into for app-only access. |

## TypeORM Definition

```typescript
@Entity('microsoft_users')
export class MicrosoftUser {
  @PrimaryGeneratedColumn('increment')
  id!: number;

  @Column({ name: 'external_user_id' })
  @Index()
  externalUserId: string = '';

  // Delegated OAuth capability (nullable for app-only-only users)
  @Column({ name: 'access_token', type: 'text', nullable: true })
  accessToken: string | null = null;

  @Column({ name: 'refresh_token', type: 'text', nullable: true })
  refreshToken: string | null = null;

  @Column({ name: 'token_expiry', type: 'datetime', nullable: true })
  tokenExpiry: Date | null = null;

  @Column({ name: 'scopes', type: 'text', nullable: true })
  scopes: string | null = null;

  @Column({ name: 'default_calendar_id', type: 'varchar', length: 255, nullable: true })
  defaultCalendarId: string | null = null;

  // App-only (tenant-wide) capability
  @ManyToOne(() => MicrosoftTenant, { nullable: true, onDelete: 'CASCADE' })
  @JoinColumn({ name: 'tenant_id' })
  tenant: MicrosoftTenant | null = null;

  @Column({ name: 'microsoft_user_id', type: 'varchar', length: 36, nullable: true })
  @Index()
  microsoftUserId: string | null = null;

  @Column({ name: 'user_principal_name', type: 'varchar', length: 320, nullable: true })
  userPrincipalName: string | null = null;

  // ... status, is_active, timestamps
}
```

## Usage

Delegated rows are created/updated automatically during the OAuth callback. The app-only
mapping columns are populated when a host registers a user for tenant-wide access:

```typescript
import { TenantUserService } from '@checkfirst/nestjs-outlook';

// Attach tenant identity to the user's row (upsert by external_user_id)
const user = await tenantUserService.registerUserMapping(
  'tenant-guid',           // Microsoft tenant ID
  'external-user-123',     // Your app's user ID
  'john@contoso.com'       // User's email (used to look up the Microsoft ID)
);

// Later, retrieve the Microsoft user ID for Graph API calls
const msUserId = await tenantUserService.getMicrosoftUserId(
  'tenant-guid',
  'external-user-123'
);
```

## Migration

- Base table created by migration `1699000000000-AddMicrosoftUserTable`.
- App-only columns (`tenant_id`, `microsoft_user_id`, `user_principal_name`) and the
  relaxation of the token columns to nullable are applied by
  `1750000000000-AddMicrosoftTenantTables`.

## Notes

- `external_user_id` is the mapping key. `registerUserMapping` upserts by it, so a user who
  has both delegated tokens and an app-only mapping is represented by **one** row — delegated
  lookups by `external_user_id` therefore stay unambiguous.
- `microsoft_user_id` is the Azure AD user object ID (immutable ID when
  `Prefer: IdType="ImmutableId"` is used).
- A row with `tenant_id`/`microsoft_user_id` set but no tokens is an app-only-only user;
  delegated token retrieval for such a user throws a clear "no delegated auth tokens" error.
- Deleting a `MicrosoftTenant` cascades to detach/remove app-only mappings referencing it.

## Used by

- [TenantUserService](tenant-user-service.md) — manages user lookup and mapping.
- [TenantCalendarService](tenant-calendar-service.md) — uses the mapping for calendar operations.
- [MicrosoftAuthService](microsoft-auth-service.md) — manages the delegated token columns.
- [DR-008: Tenant User Mapping](../decision-records/dr-008-tenant-user-mapping.md) — design rationale.
