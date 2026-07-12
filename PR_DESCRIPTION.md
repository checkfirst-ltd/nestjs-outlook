# Pull Request: Tenant-Wide Access (App-Only Authentication)

## Summary

Adds **app-only (client credentials) authentication** to enable tenant-wide calendar and user management without requiring individual user OAuth flows. This is essential for enterprise scenarios where administrators need to manage calendars for all employees in their Microsoft 365 tenant.

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                        BEFORE: Delegated Auth Only                         │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                             │
│    User A ──OAuth──► Module ──► Graph API ──► User A's Calendar             │
│    User B ──OAuth──► Module ──► Graph API ──► User B's Calendar             │
│    User C ──OAuth──► Module ──► Graph API ──► User C's Calendar             │
│    User D ──( not onboarded )──► ✗ No access                                │
│                                                                             │
│    Each user must complete OAuth. No access to non-onboarded users.         │
│                                                                             │
└─────────────────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────────────────┐
│                         AFTER: Dual Authentication                          │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                             │
│   ┌──────────────────────────────────────────────────────────────────────┐  │
│   │                    MicrosoftOutlookModule                            │  │
│   ├─────────────────────────────┬────────────────────────────────────────┤  │
│   │                             │                                        │  │
│   │   ┌─────────────────────┐   │   ┌────────────────────────────────┐   │  │
│   │   │   DELEGATED MODE    │   │   │         APP-ONLY MODE          │   │  │
│   │   │   (per-user OAuth)  │   │   │    (client credentials)        │   │  │
│   │   ├─────────────────────┤   │   ├────────────────────────────────┤   │  │
│   │   │ MicrosoftAuthService│   │   │     AppOnlyAuthService         │   │  │
│   │   │ CalendarService     │   │   │   TenantCalendarService        │   │  │
│   │   │ EmailService        │   │   │   TenantUserService            │   │  │
│   │   └─────────────────────┘   │   └────────────────────────────────┘   │  │
│   │           │                 │              │                         │  │
│   │           └────────┬────────┴──────────────┘                         │  │
│   │                    ▼                                                 │  │
│   │         ┌─────────────────────┐                                      │  │
│   │         │  Shared Components  │                                      │  │
│   │         │  - Rate limiter     │                                      │  │
│   │         │  - Retry logic      │                                      │  │
│   │         │  - Graph wrapper    │                                      │  │
│   │         └─────────────────────┘                                      │  │
│   └──────────────────────────────────────────────────────────────────────┘  │
│                                                                             │
│    App ──credentials──► Module ──► Graph API ──► ANY user's calendar        │
│                                                                             │
└─────────────────────────────────────────────────────────────────────────────┘
```

## Key Features

| Feature | Description |
|---------|-------------|
| **AppOnlyAuthService** | Client credentials flow with secret or certificate auth |
| **TenantCalendarService** | CRUD operations on any user's calendar in the tenant |
| **TenantUserService** | Enumerate users + map host app user IDs to Microsoft UPNs |
| **Identity Mapping** | Bidirectional lookup between app user IDs and Microsoft UPNs |
| **Certificate Auth** | Production-ready X.509 certificate authentication option |

## User Identity Mapping Flow

```
┌──────────────────┐       ┌─────────────────────┐       ┌──────────────────┐
│   Your App's     │       │                     │       │   Microsoft 365  │
│   User Database  │       │   TenantUserService │       │   Tenant         │
├──────────────────┤       ├─────────────────────┤       ├──────────────────┤
│                  │       │                     │       │                  │
│  user_id: "123"  │──────►│  mapUser("123",     │──────►│ Validates UPN    │
│  name: "John"    │       │   "john@co.com")    │       │ exists in tenant │
│                  │       │         │           │       │                  │
└──────────────────┘       │         ▼           │       └──────────────────┘
                           │  ┌───────────────┐  │
                           │  │ TenantUser    │  │
                           │  │ Mapping Table │  │
                           │  ├───────────────┤  │
                           │  │ external: 123 │  │
                           │  │ upn: john@... │  │
                           │  └───────────────┘  │
                           │         │           │
                           │         ▼           │
┌──────────────────┐       │  resolveUPN("123")  │       ┌──────────────────┐
│  Calendar Ops    │◄──────│  → "john@co.com"   │       │                  │
│  listEvents(upn) │       │                     │       │                  │
└──────────────────┘       └─────────────────────┘       └──────────────────┘
```

## Changes Overview

```
 src/
 ├── controllers/
 │   └── tenant-auth.controller.ts      (+446)   # Admin consent flow endpoints
 │
 ├── entities/
 │   ├── microsoft-tenant.entity.ts     (+85)    # Tenant registration entity
 │   ├── microsoft-tenant-user.entity.ts(+71)    # Tenant user cache entity
 │   └── tenant-user.entity.ts          (+82)    # User mapping entity
 │
 ├── migrations/
 │   ├── 1750000000000-AddMicrosoftTenantTables.ts
 │   ├── 1782207400000-AddTenantColumnsToSubscriptions.ts
 │   └── 1782207500000-AddTenantUsersTable.ts
 │
 ├── repositories/
 │   ├── microsoft-tenant.repository.ts (+157)
 │   └── microsoft-tenant-user.repository.ts (+219)
 │
 └── services/
     ├── auth/
     │   └── app-only-auth.service.ts   (+628)   # Client credentials token mgmt
     │
     └── tenant/
         ├── tenant-calendar.service.ts (+887)   # Tenant-wide calendar ops
         └── tenant-user.service.ts     (+548)   # User enum + mapping

 docs/
 ├── decision-records/
 │   ├── dr-006-dual-auth-architecture.md        # Why dual auth model
 │   ├── dr-007-certificate-authentication.md    # Cert auth rationale
 │   └── dr-008-tenant-user-mapping.md           # User mapping strategy
 │
 ├── how-to/
 │   └── connect-enterprise-tenant.md            # Integration guide
 │
 └── reference/
     ├── app-only-auth-service.md
     ├── tenant-calendar-service.md
     └── tenant-user-service.md

 samples/nestjs-outlook-example/
 └── src/tenant/                         # Full demo implementation
     ├── tenant.controller.ts
     ├── tenant.service.ts
     └── tenant-demo.controller.ts       # Interactive UI demo
```

## Test Coverage

| File | Lines | Tests |
|------|-------|-------|
| `app-only-auth.service.spec.ts` | 710 | Token acquisition, caching, certificate auth |
| `tenant-calendar.service.spec.ts` | 713 | CRUD, free/busy, meeting times |
| `tenant-user.service.spec.ts` | 616 | Enumeration, mapping, resolution |
| `microsoft-tenant.repository.spec.ts` | 385 | Repository operations |
| `microsoft-subscription.service.spec.ts` | 457 | Webhook subscription updates |

**Total: 3,233 lines of test code**

## Usage Example

```typescript
// Enable in module config
MicrosoftOutlookModule.forRoot({
  clientId: process.env.AZURE_CLIENT_ID,
  clientSecret: process.env.AZURE_CLIENT_SECRET,
  appOnly: {
    enabled: true,
    tenantId: process.env.AZURE_TENANT_ID,
  },
});

// Inject tenant services
@Injectable()
export class SchedulingService {
  constructor(
    private tenantCalendar: TenantCalendarService,
    private tenantUsers: TenantUserService,
  ) {}

  async bookMeeting(hostUserId: string) {
    // Resolve your app's user ID to Microsoft UPN
    const upn = await this.tenantUsers.resolveUserPrincipalName(hostUserId);

    // Create event on their calendar
    return this.tenantCalendar.createEvent(upn, {
      subject: 'Team Sync',
      start: { dateTime: '2026-06-25T10:00:00', timeZone: 'UTC' },
      end: { dateTime: '2026-06-25T11:00:00', timeZone: 'UTC' },
    });
  }
}
```

## Documentation Added

- **[DR-006: Dual Auth Architecture](docs/decision-records/dr-006-dual-auth-architecture.md)** — design rationale for coexisting auth modes
- **[DR-007: Certificate Authentication](docs/decision-records/dr-007-certificate-authentication.md)** — security rationale for cert auth
- **[DR-008: Tenant User Mapping](docs/decision-records/dr-008-tenant-user-mapping.md)** — identity mapping strategy
- **[How-To: Connect Enterprise Tenant](docs/how-to/connect-enterprise-tenant.md)** — step-by-step integration guide
- **[Reference: TenantCalendarService](docs/reference/tenant-calendar-service.md)** — full API docs
- **[Reference: TenantUserService](docs/reference/tenant-user-service.md)** — full API docs
- **[Reference: AppOnlyAuthService](docs/reference/app-only-auth-service.md)** — token service docs

## Breaking Changes

None. App-only mode is opt-in via `appOnly.enabled: true`. Existing delegated auth users are unaffected.

## Test Plan

- [ ] Run unit tests: `npm test`
- [ ] Run sample app: `cd samples/nestjs-outlook-example && npm run start:dev`
- [ ] Verify delegated auth still works (regression)
- [ ] Test app-only flow with real Azure tenant
- [ ] Test certificate auth with PEM key
- [ ] Verify user mapping persistence across restarts

---

🤖 Generated with [Claude Code](https://claude.ai/claude-code)
