---
dep:
  type: reference
  audience: [app-developer, ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/services/tenant/tenant-user.service.ts
    - src/services/auth/app-only-auth.service.ts
  tags: [users, service, tenant, app-only, mapping, api]
  links:
    - target: ./app-only-auth-service.md
      rel: USES
    - target: ./user-id-converter-service.md
      rel: RELATED
    - target: ../how-to/connect-enterprise-tenant.md
      rel: REQUIRES
    - target: ../decision-records/dr-008-tenant-user-mapping.md
      rel: EXPLAINS
---

# TenantUserService Reference

Injectable service for enumerating and mapping users in a Microsoft 365 tenant using app-only
authentication. Exported from `@checkfirst/nestjs-outlook`.

This service is only available when `appOnly.enabled` is `true` in the module configuration.
It provides two key capabilities:

1. **User enumeration:** List and search users in the Microsoft 365 tenant
2. **Identity mapping:** Map your application's user identifiers to Microsoft user principal names

## Methods

### `listUsers(options?)`

Retrieves users from the Microsoft 365 tenant.

| Parameter | Type | Description |
|-----------|------|-------------|
| `options` | `ListUsersOptions` | Filter and pagination options |

**ListUsersOptions:**

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `top` | `number` | `100` | Maximum users to return per page |
| `select` | `string[]` | core fields | Fields to include |
| `filter` | `string` | — | OData filter expression |
| `orderBy` | `string` | `'displayName'` | Sort order |
| `skipToken` | `string` | — | Pagination token from previous response |

**Returns:** `Promise<UserListResponse>`

**UserListResponse:**

| Field | Type | Description |
|-------|------|-------------|
| `users` | `TenantUser[]` | Array of user objects |
| `nextLink` | `string \| undefined` | Pagination token for next page |

**Example:**
```typescript
// List first 50 users
const result = await this.tenantUsers.listUsers({ top: 50 });

// Filter by department
const salesTeam = await this.tenantUsers.listUsers({
  filter: "department eq 'Sales'",
});

// Paginate through all users
let nextToken: string | undefined;
do {
  const page = await this.tenantUsers.listUsers({ skipToken: nextToken });
  processUsers(page.users);
  nextToken = page.nextLink;
} while (nextToken);
```

### `getUser(userIdentifier)`

Retrieves a single user by identifier.

| Parameter | Type | Description |
|-----------|------|-------------|
| `userIdentifier` | `string` | User principal name, object ID, or email |

**Returns:** `Promise<TenantUser>`

**TenantUser:**

| Field | Type | Description |
|-------|------|-------------|
| `id` | `string` | Microsoft Graph user ID (GUID) |
| `userPrincipalName` | `string` | User's UPN (typically email) |
| `displayName` | `string` | Full display name |
| `mail` | `string \| null` | Primary email address |
| `jobTitle` | `string \| null` | Job title |
| `department` | `string \| null` | Department name |
| `officeLocation` | `string \| null` | Office location |
| `mobilePhone` | `string \| null` | Mobile phone number |

### `searchUsers(query, options?)`

Searches users by display name, email, or other properties.

| Parameter | Type | Description |
|-----------|------|-------------|
| `query` | `string` | Search query string |
| `options` | `SearchUsersOptions` | Search options |

**SearchUsersOptions:**

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `top` | `number` | `25` | Maximum results |
| `searchFields` | `string[]` | `['displayName', 'mail']` | Fields to search |

**Returns:** `Promise<TenantUser[]>`

**Example:**
```typescript
const results = await this.tenantUsers.searchUsers('john', {
  top: 10,
  searchFields: ['displayName', 'mail', 'department'],
});
```

### `mapUser(externalUserId, userPrincipalName)`

Creates a mapping between your application's user identifier and a Microsoft user.

| Parameter | Type | Description |
|-----------|------|-------------|
| `externalUserId` | `string` | Your application's user identifier |
| `userPrincipalName` | `string` | Microsoft UPN to map to |

**Returns:** `Promise<void>`

**Example:**
```typescript
// When a user is provisioned in your app, map them to Microsoft
await this.tenantUsers.mapUser('app-user-123', 'john.doe@contoso.com');
```

### `unmapUser(externalUserId)`

Removes a user mapping.

| Parameter | Type | Description |
|-----------|------|-------------|
| `externalUserId` | `string` | Your application's user identifier |

**Returns:** `Promise<void>`

### `resolveUserPrincipalName(externalUserId)`

Resolves your application's user ID to a Microsoft user principal name.

| Parameter | Type | Description |
|-----------|------|-------------|
| `externalUserId` | `string` | Your application's user identifier |

**Returns:** `Promise<string>` — the mapped UPN.

**Throws:** `UserMappingNotFoundError` if no mapping exists.

**Example:**
```typescript
const upn = await this.tenantUsers.resolveUserPrincipalName('app-user-123');
// Returns: 'john.doe@contoso.com'

// Use with TenantCalendarService
const events = await this.tenantCalendar.listEvents(upn, options);
```

### `resolveExternalUserId(userPrincipalName)`

Reverse lookup: finds your application's user ID from a Microsoft UPN.

| Parameter | Type | Description |
|-----------|------|-------------|
| `userPrincipalName` | `string` | Microsoft user principal name |

**Returns:** `Promise<string>` — your application's user identifier.

**Throws:** `UserMappingNotFoundError` if no mapping exists.

### `getUserMapping(externalUserId)`

Retrieves the full mapping record for a user.

| Parameter | Type | Description |
|-----------|------|-------------|
| `externalUserId` | `string` | Your application's user identifier |

**Returns:** `Promise<UserMapping | null>`

**UserMapping:**

| Field | Type | Description |
|-------|------|-------------|
| `externalUserId` | `string` | Your application's user ID |
| `userPrincipalName` | `string` | Microsoft UPN |
| `microsoftUserId` | `string` | Microsoft Graph user ID |
| `createdAt` | `Date` | When mapping was created |
| `updatedAt` | `Date` | When mapping was last updated |

### `listMappings(options?)`

Lists all user mappings.

| Parameter | Type | Description |
|-----------|------|-------------|
| `options` | `{ skip?: number; take?: number }` | Pagination options |

**Returns:** `Promise<UserMapping[]>`

### `syncUserMappings()`

Validates all existing mappings against the tenant directory and removes stale entries.

**Returns:** `Promise<SyncResult>`

**SyncResult:**

| Field | Type | Description |
|-------|------|-------------|
| `total` | `number` | Total mappings checked |
| `valid` | `number` | Mappings still valid |
| `removed` | `number` | Stale mappings removed |
| `errors` | `string[]` | Any errors encountered |

### `clearTenantUserMappings(tenantId, options?)`

Removes a tenant's app-only footprint from the shared `microsoft_users` table during the
disconnect flow. Runs at most two bulk SQL statements (no per-row loop, no recursion). Used by
`DELETE /auth/microsoft/tenant/connection?purge=true` — see
[Disconnecting a tenant](../how-to/connect-enterprise-tenant.md#disconnecting-a-tenant).

| Parameter | Type | Description |
|-----------|------|-------------|
| `tenantId` | `string` | Azure AD tenant GUID. Matched regardless of `isActive`, so it works after deactivation. |
| `options.revokeDelegatedTokens` | `boolean` | Default `false`. When `true`, revoke each delegated refresh token at Microsoft (bounded concurrency) and delete **all** of the tenant's rows. |

**Behaviour:**

- **Default:** rows that also carry delegated OAuth tokens are unmapped (app-only columns nulled,
  delegated login preserved); pure app-only rows are deleted.
- **`revokeDelegatedTokens: true`:** delegated tokens are revoked (best-effort — failures are
  counted, never fatal), then every row for the tenant is deleted.

**Returns:** `Promise<ClearTenantMappingsResult>`

**ClearTenantMappingsResult:**

| Field | Type | Description |
|-------|------|-------------|
| `delegatedRowsUnmapped` | `number` | Dual-capability rows unmapped (app-only columns nulled). |
| `appOnlyRowsDeleted` | `number` | Rows deleted. |
| `tokensRevoked` | `number` | Delegated refresh tokens revoked at Microsoft. |
| `tokenRevocationFailures` | `number` | Revocations that failed (teardown continued regardless). |

## Caching

The service maintains an in-memory cache for user mappings:

- Mappings are cached on first access
- Cache TTL: 5 minutes (configurable)
- `mapUser` and `unmapUser` invalidate the cache
- Call `clearMappingCache()` to force refresh

## Error handling

| Error | Cause | Resolution |
|-------|-------|------------|
| `UserNotFoundError` | User identifier not found in tenant | Verify UPN or ID exists |
| `UserMappingNotFoundError` | No mapping for the external user ID | Create mapping with `mapUser` |
| `DuplicateMappingError` | UPN already mapped to different external ID | Remove existing mapping first |
| `InsufficientPermissionsError` | Missing `User.Read.All` | Grant permission in Azure AD |

## Used by

- [Connect enterprise tenant](../how-to/connect-enterprise-tenant.md) — setup and usage guide
- [TenantCalendarService](tenant-calendar-service.md) — resolves user identifiers

## Related

- [UserIdConverterService reference](user-id-converter-service.md) — delegated auth user conversion
- [AppOnlyAuthService reference](app-only-auth-service.md) — token acquisition
- [DR-008: Tenant User Mapping](../decision-records/dr-008-tenant-user-mapping.md) — design rationale
