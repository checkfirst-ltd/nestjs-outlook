---
dep:
  type: decision-record
  audience:
    - library-contributor
    - ai-agent
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/entities/microsoft-user.entity.ts
    - src/services/tenant/tenant-user.service.ts
  tags:
    - decision
    - users
    - mapping
    - identity
    - tenant
    - app-only
  links:
    - target: ../reference/tenant-user-service.md
      rel: DECIDES
    - target: ../reference/user-id-converter-service.md
      rel: RELATED
    - target: dr-006-dual-auth-architecture.md
      rel: EXTENDS
---

# DR-008: Tenant User Mapping Strategy

> **Update (2026-07-08) — storage consolidated onto `MicrosoftUser`.**
> The original decision persisted mappings in a **separate** `MicrosoftTenantUser` /
> `TenantUserMapping` table. That has been reversed: tenant mappings now live on the
> **shared `microsoft_users` table** as additional nullable columns
> (`tenant_id`, `microsoft_user_id`, `user_principal_name`), and the delegated token
> columns (`access_token`, `refresh_token`, `token_expiry`, `scopes`) are now nullable so
> an app-only user can exist without tokens.
>
> **Rationale for the reversal:**
> - **One identity, one row.** A host user is a single person whether reached via delegated
>   OAuth or app-only. Keeping one row (keyed by `externalUserId`) means shared user
>   features — default calendar, active flag, future per-user settings — are implemented and
>   migrated **once**, not per table.
> - **Same operations, different namespace.** Delegated and app-only differ only in the Graph
>   path (`/me/*` vs `/users/{upn}/*`), so the underlying identity model is the same.
> - **No collision.** Delegated code looks users up by `externalUserId` alone; a parallel
>   tenant table sharing that id would make those lookups ambiguous. `registerUserMapping`
>   therefore **upserts onto the existing row** rather than creating a second one.
>
> The mapping *contract* below (explicit mapping, bidirectional lookup, validation on map,
> no auto-sync) is unchanged — only the physical storage moved. The "Entity design" and
> "Stored in module's database" sections are superseded by this note.

## Context

In delegated authentication, the module automatically links the host application's user
identifier (`externalUserId`) to a Microsoft account when the user completes the OAuth
flow. The `MicrosoftUser` entity stores this binding, and `UserIdConverterService` provides
bidirectional lookups.

App-only authentication has no OAuth flow — the application authenticates as itself, not on
behalf of a user. When the host calls `TenantCalendarService.listEvents(userIdentifier)`,
it must specify which user's calendar to access using a Microsoft identifier (UPN or user ID).

This creates a gap: the host's user model has its own identifiers, but the module needs
Microsoft identifiers for Graph API calls. Without a mapping layer, every host application
must implement its own user-to-UPN resolution.

## Decision

Introduce `TenantUserService` with an explicit mapping layer that lets the host register
associations between its user identifiers and Microsoft UPNs:

```typescript
// Register mapping
await tenantUsers.mapUser('host-user-123', 'john.doe@contoso.com');

// Resolve when needed
const upn = await tenantUsers.resolveUserPrincipalName('host-user-123');
const events = await tenantCalendar.listEvents(upn);
```

**Key design choices:**

1. **Explicit mapping over automatic discovery:** The host explicitly calls `mapUser()`
   rather than the module guessing the association. This avoids false positives when
   email addresses don't match UPNs or when the host's user model doesn't include emails.

2. **Stored on the shared `MicrosoftUser` row:** Mappings are persisted as tenant columns
   on the existing `microsoft_users` table (one row per `externalUserId`), not a separate
   table. The host doesn't need to add columns to its own user table. See the update note
   above for why storage was consolidated.

3. **Bidirectional lookup:** Both `resolveUserPrincipalName(externalUserId)` and
   `resolveExternalUserId(upn)` are supported. Reverse lookup is useful when processing
   webhook notifications that include the Microsoft user.

4. **Validation on mapping:** When `mapUser()` is called, the service verifies the UPN
   exists in the tenant. This catches typos and invalid users early rather than at
   calendar access time.

5. **No automatic sync:** The module does not automatically discover or sync users.
   Enumeration (`listUsers()`) is provided, but the host decides when to map users.

## Entity design

The mapping columns live on the shared `microsoft_users` table (see
[MicrosoftUser entity](../reference/microsoft-user-entity.md)). A single row per host user
carries delegated tokens and/or app-only tenant identity:

```
┌───────────────────────────────────────────────────────────────┐
│                       microsoft_users                         │
├───────────────────────────────────────────────────────────────┤
│ id                  │ number (PK)                             │
│ external_user_id    │ string (indexed) — the mapping key      │
│ ── app-only (tenant) mapping columns ──                       │
│ tenant_id           │ FK → microsoft_tenants.id (nullable)    │
│ microsoft_user_id   │ string (indexed, nullable)              │
│ user_principal_name │ string (nullable)                       │
│ ── delegated OAuth columns (nullable for app-only users) ──   │
│ access_token / refresh_token / token_expiry / scopes          │
└───────────────────────────────────────────────────────────────┘
```

Lookups are keyed by `external_user_id` (indexed) with `microsoft_user_id` also indexed for
reverse resolution. `registerUserMapping` upserts by `external_user_id` so each host user
maps to exactly one row.

## Alternatives considered

### Require host to manage mapping

The module accepts only Microsoft UPNs and leaves identifier translation to the host.

**Rejected because:**
- Every host application duplicates the same boilerplate
- No standardized error handling for invalid UPNs
- Harder to document consistent patterns

### Automatic mapping via email matching

When a user is created in the host app with an email, automatically check if that email
exists as a UPN in the tenant and create the mapping.

**Rejected because:**
- Email addresses don't always match UPNs (aliases, external emails)
- Requires the host to expose its user creation lifecycle to the module
- False positives are worse than requiring explicit mapping

### Store mapping in host's database

Provide interfaces/decorators for the host to add a `microsoftUpn` column to its user entity.

**Rejected because:**
- Invasive to host's data model
- Module loses control over consistency and indexing
- TypeORM entity configuration complexity

### Real-time UPN resolution via Graph

On every operation, call Graph to look up the user by email or other attribute.

**Rejected because:**
- Adds latency to every calendar/email operation
- Rate limiting concerns with high-volume hosts
- Relies on the host having a reliable matching attribute

## Consequences

### Positive

- **Clear contract:** Host knows it must call `mapUser()` before tenant operations work
- **Early validation:** Invalid UPNs caught at mapping time, not operation time
- **Consistent performance:** Cached mappings avoid Graph lookups on hot paths
- **Bidirectional:** Webhook handlers can resolve Microsoft users back to host users

### Negative

- **Provisioning step:** Host must integrate mapping into its user provisioning flow
- **Sync burden:** If users are renamed or deleted in Microsoft, mappings become stale
- **Shared-row coupling:** Delegated and app-only data share a row, so lookups by
  `externalUserId` must tolerate rows where one capability's columns are null

### Operational

- Mappings are cached in memory with a 5-minute TTL to reduce database load
- `syncUserMappings()` can be scheduled to clean up stale mappings
- Unique indexes prevent duplicate mappings; errors are surfaced clearly

## Integration patterns

### On user creation in host app

```typescript
// When creating a user with a known Microsoft identity
await tenantUsers.mapUser(newUser.id, newUser.microsoftEmail);
```

### Bulk provisioning

```typescript
// Import users from HR system
for (const employee of hrEmployees) {
  try {
    await tenantUsers.mapUser(employee.id, employee.workEmail);
  } catch (e) {
    if (e instanceof UserNotFoundError) {
      logger.warn(`Employee ${employee.id} not found in Microsoft tenant`);
    }
  }
}
```

### Self-service mapping

```typescript
// Let users link their own Microsoft account
const upn = req.body.microsoftEmail;
const user = await tenantUsers.getUser(upn); // Validate exists
await tenantUsers.mapUser(currentUser.id, upn);
```

## Review trigger

Revisit if:
- Hosts consistently request automatic mapping via email matching
- Graph API changes affect UPN resolution patterns
- Performance analysis shows database lookups are a bottleneck
