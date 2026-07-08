---
dep:
  type: explanation
  audience:
    - library-consumer
    - library-contributor
    - ai-agent
  owner: "@checkfirst-ltd"
  created: 2026-06-25
  last_verified: 2026-06-25T12:00:00+03:00
  confidence: high
  depends_on:
    - src/services/auth/app-only-auth.service.ts
    - src/services/tenant/tenant-calendar.service.ts
    - src/services/tenant/tenant-user.service.ts
    - src/entities/microsoft-tenant.entity.ts
    - src/entities/microsoft-user.entity.ts
  tags:
    - architecture
    - authentication
    - tenant
    - app-only
    - enterprise
  links:
    - target: ../decision-records/dr-006-dual-auth-architecture.md
      rel: EXPLAINS
    - target: ../decision-records/dr-007-certificate-authentication.md
      rel: EXPLAINS
    - target: ../decision-records/dr-008-tenant-user-mapping.md
      rel: EXPLAINS
    - target: ../reference/app-only-auth-service.md
      rel: DOCUMENTS
    - target: ../reference/tenant-calendar-service.md
      rel: DOCUMENTS
    - target: ../reference/tenant-user-service.md
      rel: DOCUMENTS
    - target: ../how-to/connect-enterprise-tenant.md
      rel: NEXT
---

# Tenant-Wide Authentication

This document explains the tenant-wide (app-only) authentication architecture in depth. It
covers why two authentication modes exist, how certificate-based auth works, the user mapping
strategy, and when to use tenant-wide vs per-user authentication.

## The Big Picture

This module wraps Microsoft Graph API for NestJS apps, handling OAuth, calendars, emails, and
webhooks.

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                           YOUR NESTJS HOST APP                              │
│                                                                             │
│   ┌──────────────┐     ┌──────────────┐     ┌──────────────┐               │
│   │ Your Service │     │ Your Service │     │ @OnEvent()   │               │
│   │   (inject)   │     │   (inject)   │     │  handlers    │               │
│   └──────┬───────┘     └──────┬───────┘     └──────▲───────┘               │
│          │                    │                    │ events                 │
└──────────┼────────────────────┼────────────────────┼────────────────────────┘
           │                    │                    │
┌──────────▼────────────────────▼────────────────────┴────────────────────────┐
│                       MicrosoftOutlookModule                                │
│  ┌────────────────────────────────────────────────────────────────────────┐ │
│  │                          SERVICE LAYER                                 │ │
│  │  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────────────┐ │ │
│  │  │  CalendarService│  │  EmailService   │  │ MicrosoftAuthService    │ │ │
│  │  │  (per-user)     │  │  (per-user)     │  │ (OAuth token mgmt)      │ │ │
│  │  └─────────────────┘  └─────────────────┘  └─────────────────────────┘ │ │
│  │  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────────────┐ │ │
│  │  │TenantCalendar   │  │ TenantUser      │  │ AppOnlyAuthService      │ │ │
│  │  │Service (tenant) │  │ Service         │  │ (client credentials)    │ │ │
│  │  └─────────────────┘  └─────────────────┘  └─────────────────────────┘ │ │
│  └────────────────────────────────────────────────────────────────────────┘ │
│  ┌────────────────────────────────────────────────────────────────────────┐ │
│  │                          SHARED LAYER                                  │ │
│  │  ┌───────────────┐  ┌───────────────┐  ┌───────────────┐              │ │
│  │  │ RateLimiter   │  │ DeltaSync     │  │ UserIdConverter│              │ │
│  │  │ + CircuitBrkr │  │ Service       │  │ Service        │              │ │
│  │  └───────────────┘  └───────────────┘  └───────────────┘              │ │
│  └────────────────────────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────────────────────────┘
           │                                          ▲
           │ HTTP                                     │ Webhooks
           ▼                                          │
┌─────────────────────────────────────────────────────┴───────────────────────┐
│                        Microsoft Graph API                                  │
└─────────────────────────────────────────────────────────────────────────────┘
```

## Why Two Auth Modes?

The fundamental question is: **"Who is making the API call?"**

```
┌─────────────────────────────────────────────────────────────────┐
│                    AUTHENTICATION MODES                          │
├─────────────────────────────┬───────────────────────────────────┤
│      DELEGATED (per-user)   │        APP-ONLY (tenant-wide)     │
├─────────────────────────────┼───────────────────────────────────┤
│  MicrosoftAuthService       │  AppOnlyAuthService               │
│  ↓                          │  ↓                                │
│  User logs in → OAuth flow  │  Admin consents once → app gets   │
│  ↓                          │  client credentials grant         │
│  CalendarService            │  ↓                                │
│  EmailService               │  TenantCalendarService            │
│  (operate AS the user)      │  TenantUserService                │
│                             │  (operate ON BEHALF of users)     │
├─────────────────────────────┼───────────────────────────────────┤
│  "Read MY calendar"         │  "Read ANY user's calendar"       │
└─────────────────────────────┴───────────────────────────────────┘
```

**Delegated auth** is for scenarios where the user is present and consents to the app accessing
their data. The app acts as the user.

**App-only auth** is for enterprise scenarios where an administrator grants the app permission
to access all users' data. No individual user needs to log in.

## The App-Only Flow

### Step 1: Admin Consent (One-Time Setup)

Before the app can access tenant resources, an administrator must grant consent:

```
┌──────────────┐                    ┌──────────────┐                    ┌──────────────┐
│  Your App    │                    │   Azure AD   │                    │   Admin      │
│  (Backend)   │                    │              │                    │   (Browser)  │
└──────┬───────┘                    └──────┬───────┘                    └──────┬───────┘
       │                                   │                                   │
       │ 1. Generate admin consent URL     │                                   │
       │ ─────────────────────────────────►│                                   │
       │                                   │                                   │
       │ URL: https://login.microsoftonline.com/{tenant}/adminconsent          │
       │      ?client_id=...&redirect_uri=...                                  │
       │                                   │                                   │
       │◄──────────────────────────────────│                                   │
       │                                   │                                   │
       │                                   │ 2. Admin visits URL               │
       │                                   │◄──────────────────────────────────│
       │                                   │                                   │
       │                                   │ 3. Admin sees permission list:    │
       │                                   │    - Read all users' calendars    │
       │                                   │    - Read/write all users' mail   │
       │                                   │    - Read directory data          │
       │                                   │──────────────────────────────────►│
       │                                   │                                   │
       │                                   │ 4. Admin clicks "Accept"          │
       │                                   │◄──────────────────────────────────│
       │                                   │                                   │
       │ 5. Redirect with admin_consent=True, tenant=GUID                      │
       │◄──────────────────────────────────│                                   │
       │                                   │                                   │
       │ 6. Store tenant ID in database    │                                   │
       │    (MicrosoftTenant entity)       │                                   │
       │                                   │                                   │
```

The `AppOnlyAuthService` provides methods for this flow:

```typescript
// Generate the URL for admin to visit
const consentUrl = appOnlyAuthService.getAdminConsentUrl(state, 'common');
// → https://login.microsoftonline.com/common/adminconsent?client_id=...

// Handle the callback when admin is redirected back
const result = appOnlyAuthService.handleAdminConsentCallback(queryParams);
// → { tenantId: 'contoso-guid', success: true }
```

### Step 2: Getting Access Tokens (Certificate Auth)

Once consent is granted, the app can request tokens. Certificate authentication is recommended
for production because the private key never leaves your server.

#### Why Certificates Instead of Secrets?

```
┌──────────────────────────────────────────────────────────────────────────────┐
│                    WHY CERTIFICATES INSTEAD OF SECRETS?                      │
├──────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│  CLIENT SECRET (simpler, less secure):                                       │
│  ┌──────────┐                      ┌──────────┐                              │
│  │ Your App │ ───── secret ──────► │ Azure AD │                              │
│  └──────────┘   "abc123xyz..."     └──────────┘                              │
│                                                                              │
│  Problem: The secret travels over the network. If intercepted, attacker      │
│           can impersonate your app indefinitely.                             │
│                                                                              │
├──────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│  CERTIFICATE (enterprise-grade):                                             │
│  ┌──────────┐                      ┌──────────┐                              │
│  │ Your App │ ─── signed JWT ────► │ Azure AD │                              │
│  └──────────┘                      └──────────┘                              │
│       │                                  │                                   │
│       │ Private key stays HERE           │ Has public key                    │
│       │ (never transmitted)              │ (verifies signature)              │
│                                                                              │
│  Advantage: Even if someone intercepts the JWT, it expires in 10 minutes    │
│             and they can't create new ones without the private key.          │
│                                                                              │
└──────────────────────────────────────────────────────────────────────────────┘
```

#### The Token Request Flow

```
┌───────────────────────────────────────────────────────────────────────────────┐
│                         CLIENT CREDENTIALS FLOW                               │
└───────────────────────────────────────────────────────────────────────────────┘

┌──────────────┐                                           ┌──────────────────┐
│  Your App    │                                           │  Azure AD        │
│              │                                           │  Token Endpoint  │
└──────┬───────┘                                           └────────┬─────────┘
       │                                                            │
       │  1. Build JWT Client Assertion                             │
       │     ┌─────────────────────────────────────────┐            │
       │     │ HEADER:                                 │            │
       │     │   alg: "PS256"                          │            │
       │     │   typ: "JWT"                            │            │
       │     │   x5t#S256: "<cert thumbprint>"         │◄── identifies which cert
       │     ├─────────────────────────────────────────┤            │
       │     │ PAYLOAD:                                │            │
       │     │   iss: "your-client-id"                 │            │
       │     │   sub: "your-client-id"                 │            │
       │     │   aud: "https://login.../token"         │            │
       │     │   jti: "<random-uuid>"                  │◄── prevents replay
       │     │   exp: <now + 10 minutes>               │            │
       │     │   iat: <now>                            │            │
       │     ├─────────────────────────────────────────┤            │
       │     │ SIGNATURE:                              │            │
       │     │   RSA-PSS-SHA256(header.payload, key)   │            │
       │     └─────────────────────────────────────────┘            │
       │                                                            │
       │  2. POST to token endpoint                                 │
       │     grant_type=client_credentials                          │
       │     client_id=...                                          │
       │     scope=https://graph.microsoft.com/.default             │
       │     client_assertion_type=urn:ietf:...:jwt-bearer          │
       │     client_assertion=<the JWT above>                       │
       │ ──────────────────────────────────────────────────────────►│
       │                                                            │
       │                                        3. Verify signature │
       │                                           Check thumbprint │
       │                                           Validate claims  │
       │                                                            │
       │  4. Return access_token (valid ~1 hour)                    │
       │ ◄──────────────────────────────────────────────────────────│
       │                                                            │
       │  5. Cache token until 5 min before expiry                  │
       │                                                            │
```

The `AppOnlyAuthService` handles all of this internally:

```typescript
// Get a token - handles caching, JWT signing, everything
const token = await appOnlyAuthService.getAccessToken(tenantId);
// or
const token = await appOnlyAuthService.getAccessToken(microsoftTenantEntity);
```

## The Data Model

Tenant registration lives in `microsoft_tenants`. User mappings do **not** have their own
table — they are stored as nullable columns on the shared `microsoft_users` table, so one
row per host user carries delegated tokens and/or app-only tenant identity (see
[DR-008](../decision-records/dr-008-tenant-user-mapping.md)):

```
┌────────────────────────────────┐         ┌────────────────────────────────────┐
│     microsoft_tenants          │         │          microsoft_users           │
├────────────────────────────────┤         ├────────────────────────────────────┤
│ id (PK)                        │────┐    │ id (PK)                            │
│ tenant_id (unique)             │    │    │ external_user_id (indexed)         │
│ client_id                      │    │    │ ── delegated (nullable) ──         │
│ certificate_thumbprint         │    │    │ access_token / refresh_token       │
│ certificate_path               │    │    │ token_expiry / scopes / status     │
│ certificate_key_path           │    │    │ ── app-only (nullable) ──          │
│ status                         │    └───►│ tenant_id (FK)                     │
│ admin_consent_granted_at       │         │ microsoft_user_id (indexed)        │
│ is_active                      │         │ user_principal_name                │
│ created_at                     │         │ default_calendar_id (cached)       │
│ updated_at                     │         │ is_active / created_at / updated_at│
└────────────────────────────────┘         └────────────────────────────────────┘

Status values:
  PENDING_CONSENT  →  Admin hasn't granted consent yet
  ACTIVE           →  Ready to use
  CONSENT_REVOKED  →  Admin revoked permissions
  ERROR            →  Something went wrong
```

## User Mapping

The key challenge: **Your app has its own user IDs. Microsoft has different user IDs.**

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                    THE USER IDENTITY PROBLEM                                │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                             │
│  YOUR APP (ScheduleAI)              MICROSOFT GRAPH                         │
│  ──────────────────────             ───────────────                         │
│                                                                             │
│  Inspector ID: "insp-12345"    ←→   Microsoft User ID: "abc-def-ghi-789"   │
│  Name: "John Doe"                   UPN: "john.doe@contoso.com"             │
│                                                                             │
│  How do we know they're the same person?                                    │
│                                                                             │
└─────────────────────────────────────────────────────────────────────────────┘
```

**Solution:** Explicit mapping via `TenantUserService`.

```
┌──────────────┐                    ┌──────────────────┐
│  Your App    │                    │ TenantUserService│
│  (onboarding │                    │                  │
│   an insp.)  │                    │                  │
└──────┬───────┘                    └────────┬─────────┘
       │                                     │
       │ 1. registerUserMapping(             │
       │      tenantId: "contoso-tenant",    │
       │      externalUserId: "insp-12345",  │
       │      email: "john.doe@contoso.com"  │
       │    )                                │
       │ ───────────────────────────────────►│
       │                                     │
       │                                     │ 2. Call Graph API:
       │                                     │    GET /users?$filter=mail eq '...'
       │                                     │    or userPrincipalName eq '...'
       │                                     │
       │                                     │    → Returns microsoftUserId
       │                                     │
       │                                     │ 3. Upsert mapping onto microsoft_users
       │                                     │    (by external_user_id):
       │                                     │    ┌─────────────────────────────┐
       │                                     │    │ external: "insp-12345"      │
       │                                     │    │ microsoft: "abc-def-789"    │
       │                                     │    │ upn: "john.doe@contoso.com" │
       │                                     │    └─────────────────────────────┘
       │                                     │
       │◄────────────────────────────────────│
       │  Returns MicrosoftUser entity       │
       │                                     │
```

Later, when you need to access their calendar:

```
┌──────────────┐         ┌───────────────────┐         ┌─────────────────────┐
│  Your App    │         │ TenantUserService │         │ TenantCalendarService│
└──────┬───────┘         └─────────┬─────────┘         └──────────┬──────────┘
       │                           │                              │
       │ getMicrosoftUserId(       │                              │
       │   "contoso-tenant",       │                              │
       │   "insp-12345"            │                              │
       │ )                         │                              │
       │ ─────────────────────────►│                              │
       │                           │                              │
       │   "abc-def-789"           │                              │
       │ ◄─────────────────────────│                              │
       │                           │                              │
       │ createEvent(event, "contoso-tenant", "abc-def-789", ...)│
       │ ────────────────────────────────────────────────────────►│
       │                                                          │
       │                          POST /users/abc-def-789/calendars/.../events
       │                                                          │
```

## TenantCalendarService vs CalendarService

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                        SERVICE COMPARISON                                   │
├────────────────────────────────┬────────────────────────────────────────────┤
│       CalendarService          │        TenantCalendarService               │
│       (delegated auth)         │        (app-only auth)                     │
├────────────────────────────────┼────────────────────────────────────────────┤
│                                │                                            │
│  // Inject into constructor    │  // Inject into constructor                │
│  CalendarService               │  TenantCalendarService                     │
│                                │                                            │
│  // Read events                │  // Read events                            │
│  getEvents(                    │  getEventById(                             │
│    externalUserId: string      │    tenantId: string,                       │
│  )                             │    microsoftUserId: string,                │
│                                │    eventId: string                         │
│  Uses: /me/events              │  )                                         │
│  Token: User's OAuth token     │  Uses: /users/{id}/events/{id}             │
│                                │  Token: App-only token                     │
│                                │                                            │
│  // Create event               │  // Create event                           │
│  createEvent(                  │  createEvent(                              │
│    event,                      │    event,                                  │
│    externalUserId              │    tenantId,                               │
│  )                             │    microsoftUserId,                        │
│                                │    calendarId                              │
│  Uses: /me/calendars/.../events│  )                                         │
│                                │  Uses: /users/{id}/calendars/.../events    │
│                                │                                            │
├────────────────────────────────┼────────────────────────────────────────────┤
│  Token management:             │  Token management:                         │
│  Per-user refresh tokens       │  Single tenant-wide token (cached)         │
│  stored in MicrosoftUser       │  refreshed automatically                   │
│  entity                        │                                            │
│                                │                                            │
│  Token refresh when expired    │  Token requested via client credentials    │
│  using user's refresh_token    │  (no user interaction needed)              │
│                                │                                            │
└────────────────────────────────┴────────────────────────────────────────────┘
```

## Batch Operations

`TenantCalendarService` supports batch operations for efficiency:

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                    BATCH API ($batch endpoint)                              │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                             │
│  INSTEAD OF:                        DO THIS:                                │
│                                                                             │
│  POST /users/123/calendars/.../events   POST /$batch                        │
│  POST /users/123/calendars/.../events   {                                   │
│  POST /users/123/calendars/.../events     "requests": [                     │
│  POST /users/123/calendars/.../events       { id: "0", method: "POST", ... }│
│  POST /users/123/calendars/.../events       { id: "1", method: "POST", ... }│
│  ...                                        { id: "2", method: "POST", ... }│
│  (20 separate HTTP calls)                   ...                             │
│                                           ]                                 │
│                                         }                                   │
│                                         (1 HTTP call, up to 20 ops)         │
│                                                                             │
├─────────────────────────────────────────────────────────────────────────────┤
│  Methods available:                                                         │
│  - createBatchEvents()   →  Create up to 20 events at once                  │
│  - updateBatchEvents()   →  Update up to 20 events at once                  │
│  - deleteBatchEvents()   →  Delete up to 20 events at once                  │
│  - getEventsBatch()      →  Fetch up to 20 events at once                   │
│                                                                             │
│  Results include per-item success/failure:                                  │
│  [                                                                          │
│    { index: 0, success: true, event: {...} },                               │
│    { index: 1, success: false, error: "Conflict" },                         │
│    { index: 2, success: true, event: {...} },                               │
│  ]                                                                          │
│                                                                             │
└─────────────────────────────────────────────────────────────────────────────┘
```

## Complete Flow Example

```
┌─────────────────────────────────────────────────────────────────────────────┐
│              SCHEDULING AN INSPECTION IN CONTOSO'S TENANT                   │
└─────────────────────────────────────────────────────────────────────────────┘

SETUP PHASE (once per tenant):
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. Contoso admin visits admin consent URL
2. Admin grants permissions (Calendars.ReadWrite, User.Read.All)
3. Store MicrosoftTenant entity:
   ┌─────────────────────────────────────────┐
   │ tenantId: "contoso-guid"                │
   │ clientId: "your-app-guid"               │
   │ certificateThumbprint: "abc123..."      │
   │ certificateKeyPath: "/certs/contoso.key"│
   │ status: ACTIVE                          │
   └─────────────────────────────────────────┘

4. Map inspectors to Microsoft users:
   registerUserMapping("contoso-guid", "insp-001", "alice@contoso.com")
   registerUserMapping("contoso-guid", "insp-002", "bob@contoso.com")


RUNTIME PHASE (for each scheduling operation):
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Your scheduler decides: "Alice needs to inspect Site X tomorrow at 2pm"

1. Resolve Alice's Microsoft ID:
   microsoftUserId = getMicrosoftUserId("contoso-guid", "insp-001")
   → Returns "alice-microsoft-guid"

2. Get Alice's default calendar:
   calendarId = getDefaultCalendarId("contoso-guid", "alice-microsoft-guid")
   → Returns "AAMkAGI2TG..."

3. Create the event:
   createEvent(
     {
       subject: "Site X Inspection",
       start: { dateTime: "2024-03-15T14:00:00", timeZone: "UTC" },
       end: { dateTime: "2024-03-15T16:00:00", timeZone: "UTC" },
       location: { displayName: "Site X, Building 3" },
     },
     "contoso-guid",           // tenantId
     "alice-microsoft-guid",   // microsoftUserId
     "AAMkAGI2TG..."           // calendarId
   )

4. Under the hood:
   - AppOnlyAuthService.getAccessToken("contoso-guid")
     → Checks cache, or builds JWT assertion + requests new token
   - POST to Graph API: /users/{alice}/calendars/{cal}/events
   - Event created in Alice's Outlook calendar!
```

## Configuration Options

```typescript
// In your NestJS module setup:
MicrosoftOutlookModule.forRoot({
  clientId: 'your-app-guid',
  clientSecret: 'for-delegated-auth',     // Still needed for per-user auth

  // App-only configuration
  appOnly: {
    enabled: true,
    tenantId: 'your-default-tenant',      // Optional: single-tenant mode

    // Certificate auth (recommended)
    certificate: {
      thumbprint: 'SHA256-OF-YOUR-CERT',
      // Choose ONE of these:
      privateKey: '-----BEGIN PRIVATE KEY-----\n...',   // Direct string
      privateKeyPath: '/etc/secrets/app.key',           // File path
      privateKeyBase64: 'LS0tLS1CRUdJTi4uLg==',        // Base64 encoded
    },

    // Optional: custom scopes
    scopes: ['https://graph.microsoft.com/.default'],

    // Optional: token cache TTL (default: 55 min)
    tokenCacheTtlMs: 55 * 60 * 1000,
  },

  // ... other config
});
```

## When to Use What

| Scenario | Use | Why |
|----------|-----|-----|
| User connects their personal calendar | `CalendarService` (delegated) | User must consent, app acts as user |
| Scheduling system for enterprise | `TenantCalendarService` (app-only) | Admin consents once, app manages all |
| User manages their own email | `EmailService` (delegated) | Personal email access |
| Background sync of all calendars | `TenantCalendarService` (app-only) | No user present, bulk operations |
| Single-tenant app (your company only) | App-only with config credentials | Simple setup |
| Multi-tenant SaaS | App-only with `MicrosoftTenant` entities | Per-tenant credentials |

## Related Documentation

- [DR-006: Dual Authentication Architecture](../decision-records/dr-006-dual-auth-architecture.md) — the decision behind supporting both modes
- [DR-007: Certificate Authentication](../decision-records/dr-007-certificate-authentication.md) — why certificates over secrets
- [DR-008: Tenant User Mapping](../decision-records/dr-008-tenant-user-mapping.md) — the user mapping strategy
- [How to: Connect an Enterprise Tenant](../how-to/connect-enterprise-tenant.md) — step-by-step guide
- [Reference: AppOnlyAuthService](../reference/app-only-auth-service.md) — full API documentation
- [Reference: TenantCalendarService](../reference/tenant-calendar-service.md) — calendar operations API
- [Reference: TenantUserService](../reference/tenant-user-service.md) — user mapping API
