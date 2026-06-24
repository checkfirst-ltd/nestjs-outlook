---
dep:
  type: reference
  audience: [app-developer, ai-agent, library-contributor]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/services/auth/app-only-auth.service.ts
    - src/interfaces/config/outlook-config.interface.ts
  tags: [auth, service, app-only, client-credentials, tenant, api]
  links:
    - target: ./configuration.md
      rel: USES
    - target: ../how-to/connect-enterprise-tenant.md
      rel: NEXT
    - target: ../decision-records/dr-006-dual-auth-architecture.md
      rel: EXPLAINS
    - target: ../decision-records/dr-007-certificate-authentication.md
      rel: EXPLAINS
---

# AppOnlyAuthService Reference

Injectable service that handles the Microsoft OAuth 2.0 client credentials flow for app-only
(tenant-wide) authentication. Exported from `@checkfirst/nestjs-outlook`.

This service is only available when `appOnly.enabled` is `true` in the module configuration.
It authenticates the application itself rather than on behalf of a user, enabling access to
all resources in the configured tenant.

## Methods

### `getAccessToken()`

Acquires or returns a cached access token for the configured tenant using the client
credentials flow.

**Returns:** `Promise<string>` — a valid access token for Microsoft Graph API calls.

**Behavior:**
- Returns cached token if valid (not within 5 minutes of expiry)
- Acquires new token using client credentials if cache miss or expired
- Uses certificate authentication if configured, otherwise client secret
- Throws if token acquisition fails

**Example:**
```typescript
const token = await this.appOnlyAuthService.getAccessToken();
// Use token for Graph API calls
```

### `getAccessTokenForResource(resource?)`

Acquires an access token for a specific resource endpoint.

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `resource` | `string` | `'https://graph.microsoft.com'` | The resource URI to acquire a token for. |

**Returns:** `Promise<string>` — access token scoped to the specified resource.

**Example:**
```typescript
// Token for Graph API (default)
const graphToken = await this.appOnlyAuthService.getAccessTokenForResource();

// Token for a different Microsoft resource
const otherToken = await this.appOnlyAuthService.getAccessTokenForResource(
  'https://management.azure.com'
);
```

### `isConfigured()`

Checks whether app-only authentication is properly configured and available.

**Returns:** `boolean` — `true` if `appOnly.enabled` is `true` and required credentials
are present.

**Example:**
```typescript
if (this.appOnlyAuthService.isConfigured()) {
  // App-only features are available
}
```

### `clearTokenCache()`

Clears the cached access token, forcing the next `getAccessToken()` call to acquire a
fresh token.

**Returns:** `void`

**Use cases:**
- After rotating credentials (client secret or certificate)
- During testing to verify token acquisition
- When troubleshooting authentication issues

## Properties

### `tenantId`

**Type:** `string` (read-only)

The Azure AD tenant ID this service is configured for.

## Token caching

The service caches access tokens in memory to minimize token requests:

- Default TTL: 55 minutes (tokens typically expire after 60 minutes)
- Configurable via `appOnly.tokenCacheTtlMs` in module config
- Cache is process-local; each container maintains its own cache
- Token is proactively refreshed 5 minutes before expiry

## Authentication methods

The service supports two authentication methods, selected based on configuration:

### Client secret

Used when `clientSecret` is provided and no certificate is configured.

```typescript
appOnly: {
  enabled: true,
  tenantId: 'your-tenant-id',
  // Uses clientSecret from parent config
}
```

### Certificate

Used when `appOnly.certificate` is configured. Takes precedence over client secret.

```typescript
appOnly: {
  enabled: true,
  tenantId: 'your-tenant-id',
  certificate: {
    thumbprint: 'CERT_THUMBPRINT',
    privateKey: '-----BEGIN PRIVATE KEY-----\n...',
  },
}
```

See [DR-007: Certificate Authentication](../decision-records/dr-007-certificate-authentication.md)
for security considerations.

## Error handling

| Error | Cause | Resolution |
|-------|-------|------------|
| `AppOnlyAuthNotConfiguredError` | `appOnly.enabled` is `false` or missing | Enable app-only auth in config |
| `TenantAuthenticationError` | Client credentials rejected | Verify client ID, secret/certificate, and tenant ID |
| `InsufficientPermissionsError` | Missing Graph API permissions | Grant required Application permissions in Azure AD |

## Used by

- [TenantCalendarService](tenant-calendar-service.md) — acquires tokens for calendar operations
- [TenantUserService](tenant-user-service.md) — acquires tokens for user operations
- [Connect enterprise tenant](../how-to/connect-enterprise-tenant.md) — setup guide

## Related

- [Configuration reference](configuration.md) — `AppOnlyAuthConfig` interface
- [MicrosoftAuthService](microsoft-auth-service.md) — delegated (per-user) authentication
