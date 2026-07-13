---
dep:
  type: how-to
  audience: [library-integrator, app-developer]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-28T12:00:00+03:00
  confidence: high
  depends_on:
    - src/interfaces/config/outlook-config.interface.ts
    - src/services/auth/app-only-auth.service.ts
    - src/controllers/tenant-auth.controller.ts
    - src/repositories/microsoft-tenant.repository.ts
  tags: [auth, enterprise, tenant, app-only, client-credentials]
  links:
    - target: ../reference/app-only-auth-service.md
      rel: USES
    - target: ../reference/tenant-calendar-service.md
      rel: USES
    - target: ../reference/tenant-user-service.md
      rel: USES
    - target: ../reference/configuration.md
      rel: REQUIRES
    - target: ../decision-records/dr-006-dual-auth-architecture.md
      rel: EXPLAINS
---

# Connect an Enterprise Tenant

**Goal:** Configure the module for app-only (client credentials) authentication to access
calendars and emails for all users in a Microsoft 365 tenant without individual user consent.

## When to use app-only authentication

App-only authentication is appropriate when:

- Your application needs to access resources for multiple users in an organization
- You cannot or do not want to require each user to complete an OAuth flow
- An Azure AD administrator can grant tenant-wide consent
- You need background/daemon access without user interaction

For individual user authentication with delegated permissions, see
[Authenticate a User](authenticate-a-user.md) instead.

## Prerequisites

1. An Azure AD application registration with:
   - A client secret or certificate configured
   - Application permissions (not delegated) for the required Microsoft Graph APIs
   - Admin consent granted for the tenant

2. Required Microsoft Graph **Application** permissions:
   - `Calendars.ReadWrite` — read/write calendars for all users
   - `Mail.ReadWrite` — read/write mail for all users (if using email features)
   - `User.Read.All` — read user profiles to enumerate tenant users

## Steps

### 1. Configure Azure AD application

In the Azure Portal, navigate to your app registration:

1. Go to **API permissions** > **Add a permission** > **Microsoft Graph** > **Application permissions**
2. Add the required permissions listed above
3. Click **Grant admin consent for [Your Organization]**
4. Go to **Certificates & secrets** and create either:
   - A client secret (simpler, but must be rotated)
   - A certificate (more secure, recommended for production)

### 2. Register the module with app-only config

```typescript
import { MicrosoftOutlookModule } from '@checkfirst/nestjs-outlook';

@Module({
  imports: [
    MicrosoftOutlookModule.forRoot({
      clientId: process.env.AZURE_CLIENT_ID,
      clientSecret: process.env.AZURE_CLIENT_SECRET,
      redirectPath: 'auth/microsoft/callback', // Still needed for hybrid scenarios
      backendBaseUrl: process.env.BACKEND_URL,

      // Enable app-only authentication
      appOnly: {
        enabled: true,
        tenantId: process.env.AZURE_TENANT_ID, // Your organization's tenant ID

        // Optional: custom token cache TTL (default: 55 minutes)
        tokenCacheTtlMs: 3300000,
      },
    }),
  ],
})
export class AppModule {}
```

### 3. Access tenant-wide services

With app-only auth enabled, inject the tenant services instead of user-scoped services:

```typescript
import { Injectable } from '@nestjs/common';
import {
  TenantCalendarService,
  TenantUserService
} from '@checkfirst/nestjs-outlook';

@Injectable()
export class SchedulingService {
  constructor(
    private readonly tenantCalendar: TenantCalendarService,
    private readonly tenantUsers: TenantUserService,
  ) {}

  async getEmployeeCalendar(userPrincipalName: string) {
    // Access any user's calendar in the tenant
    const events = await this.tenantCalendar.listEvents(userPrincipalName, {
      startDateTime: new Date(),
      endDateTime: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000),
    });
    return events;
  }

  async listTenantUsers() {
    // Enumerate users in the tenant
    return this.tenantUsers.listUsers();
  }
}
```

### 4. Map external users to Microsoft identities

The `TenantUserService` provides methods to resolve your application's user IDs to Microsoft
user principal names:

```typescript
// Register a mapping between your user ID and the Microsoft UPN
await this.tenantUsers.mapUser('your-user-123', 'john.doe@contoso.com');

// Later, resolve the UPN from your user ID
const upn = await this.tenantUsers.resolveUserPrincipalName('your-user-123');
const events = await this.tenantCalendar.listEvents(upn);
```

## Multi-tenant: register tenants and run admin consent

The configuration above describes the **single-tenant** mode, where one tenant is
fixed in module config (`appOnly.tenantId`). For **multi-tenant** scenarios the module
also supports an **entity-based** flow, where each tenant is stored as a
`MicrosoftTenant` row and activated through the admin-consent callback handled by
`TenantAuthController` (`GET /auth/microsoft/tenant/admin-callback`).

> **Important:** the admin-consent callback only *activates* an existing
> `MicrosoftTenant` row — it never creates one. You must register the tenant first,
> otherwise the callback responds with *"Tenant Not Found — please ensure the tenant
> was properly registered before requesting admin consent."*
>
> The callback looks the tenant up by the `state` query parameter, and this
> implementation maps `state` directly to `tenantId`. So the value you store in
> `tenantId` (the Azure AD **directory GUID**) must also be the `state` you pass to
> the admin-consent URL.

### 1. Register the tenant

Insert a `MicrosoftTenant` row before initiating consent. The
[`AppOnlyAuthService.getAdminConsentUrl(state, tenantId, clientId)`](../reference/app-only-auth-service.md)
helper builds the URL to hand to the administrator:

```typescript
// state === tenantId === the Azure AD directory GUID
const tenant = await tenantRepository.save({
  tenantId: directoryGuid,          // Azure AD directory ID (GUID)
  clientId: azureAppClientId,
  certificateThumbprint: certThumbprint, // '' is allowed for client-secret auth
  status: MicrosoftTenantStatus.PENDING_CONSENT,
  isActive: true,
});

const adminConsentUrl = appOnlyAuthService.getAdminConsentUrl(
  tenant.tenantId, // state
  tenant.tenantId, // tenantId (directory the admin signs into)
  tenant.clientId,
);
```

The sample app exposes this as `POST /tenant/register` (see
`samples/nestjs-outlook-example/src/tenant/tenant.service.ts`), which returns the
ready-to-use `adminConsentUrl` in its response.

### 2. Administrator grants consent

Direct the tenant administrator to the `adminConsentUrl`. After they approve,
Microsoft redirects to `/auth/microsoft/tenant/admin-callback` with
`tenant`, `state`, and `admin_consent=True`. The callback matches the row by
`state == tenantId`, sets its status to `ACTIVE`, and verifies it can acquire a
token. The tenant is then ready for `TenantCalendarService` / `TenantUserService`.

## Disconnecting a tenant

`DELETE /auth/microsoft/tenant/connection` tears down a tenant connection. It has two modes,
controlled by query flags:

| Query | Effect | Response |
|-------|--------|----------|
| *(none)* | **Soft disconnect** (synchronous). Flags the `MicrosoftTenant` row inactive (`is_active = false`) and drops the cached app-only token. Mapped `microsoft_users` rows and Outlook webhook subscriptions are **left intact** — re-consent reactivates the tenant and existing mappings keep working. | `200` |
| `?purge=true` | **Full teardown** (runs in the background). Additionally deletes the tenant's Outlook webhook subscriptions at Microsoft (via `$batch`, 20 per call) and clears its user mappings: rows that also hold delegated OAuth tokens are unmapped (app-only columns nulled, delegated login preserved), pure app-only rows are deleted — both via bulk SQL. | `202` |
| `?revokeUserTokens=true` | Implies `purge`. Also revokes each delegated refresh token at Microsoft (bounded concurrency) and deletes **all** of the tenant's rows. | `202` |

```bash
# Soft disconnect (default) — keeps mappings and subscriptions (200, synchronous)
curl -X DELETE '.../auth/microsoft/tenant/connection?tenantId=<guid>'

# Full teardown — remove subscriptions + user mappings (202, runs in background)
curl -X DELETE '.../auth/microsoft/tenant/connection?tenantId=<guid>&purge=true'

# Full teardown + revoke delegated user tokens (202)
curl -X DELETE '.../auth/microsoft/tenant/connection?tenantId=<guid>&purge=true&revokeUserTokens=true'
```

**Why a purge is asynchronous.** A tenant may have thousands of subscriptions and users;
tearing everything down inline would exceed the request timeout and tie up a worker. The purge
therefore returns `202 Accepted` immediately and runs in the background. Inside the teardown the
tenant is deactivated **last**, so subscription deletion still has a valid app-only token while
it runs, and **the connection reporting as disconnected (`GET connection` returns "not
connected") is the completion signal** — poll it to know when a purge has finished. A `404` from
Microsoft (subscription already gone) is treated as success. `tenantId` still defaults to the
module-configured tenant when omitted.

> **Scale note:** subscription deletion is batched (`$batch`, ≤20/call) and local rows are
> deactivated in a single bulk `UPDATE`, so the work is roughly O(N/20) Graph round-trips plus a
> handful of set-based SQL statements — it does not issue one request or one `UPDATE` per user.

## Using certificate authentication

For production environments, certificate authentication is more secure than client secrets:

```typescript
import * as fs from 'fs';

MicrosoftOutlookModule.forRoot({
  clientId: process.env.AZURE_CLIENT_ID,
  clientSecret: '', // Not used with certificate auth
  redirectPath: 'auth/microsoft/callback',
  backendBaseUrl: process.env.BACKEND_URL,

  appOnly: {
    enabled: true,
    tenantId: process.env.AZURE_TENANT_ID,

    // Certificate authentication
    certificate: {
      thumbprint: process.env.AZURE_CERT_THUMBPRINT,
      privateKey: fs.readFileSync('/path/to/private-key.pem', 'utf8'),
    },
  },
}),
```

See [DR-007: Certificate Authentication](../decision-records/dr-007-certificate-authentication.md)
for the security rationale.

## Verify

- Confirm `TenantCalendarService` and `TenantUserService` are injectable
- Call `tenantUsers.listUsers()` and verify it returns users from the tenant
- Call `tenantCalendar.listEvents('user@yourdomain.com')` for a known user
- Check logs for successful token acquisition with client credentials flow

## Related

- [AppOnlyAuthService reference](../reference/app-only-auth-service.md) — token acquisition details
- [TenantCalendarService reference](../reference/tenant-calendar-service.md) — calendar operations
- [TenantUserService reference](../reference/tenant-user-service.md) — user enumeration and mapping
- [DR-006: Dual Auth Architecture](../decision-records/dr-006-dual-auth-architecture.md) — design rationale
- [Configuration reference](../reference/configuration.md) — full config options
