---
dep:
  type: how-to
  audience: [library-integrator, app-developer]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/interfaces/config/outlook-config.interface.ts
    - src/services/auth/app-only-auth.service.ts
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
