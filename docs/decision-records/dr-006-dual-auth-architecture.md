---
dep:
  type: decision-record
  audience: [library-contributor, ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/interfaces/config/outlook-config.interface.ts
    - src/services/auth/microsoft-auth.service.ts
    - src/services/auth/app-only-auth.service.ts
  tags: [decision, auth, architecture, app-only, delegated, dual-mode]
  links:
    - target: ../reference/app-only-auth-service.md
      rel: DECIDES
    - target: ../reference/microsoft-auth-service.md
      rel: DECIDES
    - target: ../how-to/connect-enterprise-tenant.md
      rel: EXPLAINS
    - target: ../explanation/architecture-overview.md
      rel: EXTENDS
---

# DR-006: Dual Authentication Architecture

## Context

The module originally supported only delegated (per-user) authentication where each user
completes an OAuth flow and the application acts on their behalf. Enterprise customers
requested tenant-wide access to manage calendars for all employees without requiring each
user to authenticate individually. Microsoft Graph supports this via the OAuth 2.0 client
credentials flow with application permissions.

The challenge is how to add app-only authentication without:
- Breaking existing delegated auth users
- Duplicating service logic for calendar/email operations
- Creating configuration complexity
- Introducing security regressions

## Decision

Implement a **dual authentication architecture** where delegated and app-only modes coexist:

1. **Separate authentication services:** `MicrosoftAuthService` handles delegated auth,
   `AppOnlyAuthService` handles client credentials. They share no state.

2. **Parallel service layers:** Tenant-scoped services (`TenantCalendarService`,
   `TenantUserService`) wrap app-only auth, while user-scoped services (`CalendarService`,
   `EmailService`) continue using delegated auth. Both call the same underlying Graph
   operations.

3. **Configuration-driven activation:** App-only mode is opt-in via `appOnly.enabled: true`.
   The tenant services are only registered when enabled; attempting to inject them without
   configuration throws a clear error.

4. **User identity mapping:** Since app-only mode lacks the user-to-Microsoft binding that
   delegated auth provides, a mapping layer lets the host associate its user IDs with
   Microsoft UPNs.

```
┌────────────────────────────────────────────────────────────────┐
│                     MicrosoftOutlookModule                     │
├────────────────────────────────────────────────────────────────┤
│                                                                │
│   ┌──────────────────────┐    ┌──────────────────────┐        │
│   │   Delegated Mode     │    │    App-Only Mode     │        │
│   │  (per-user OAuth)    │    │ (client credentials) │        │
│   ├──────────────────────┤    ├──────────────────────┤        │
│   │ MicrosoftAuthService │    │  AppOnlyAuthService  │        │
│   │   CalendarService    │    │ TenantCalendarService│        │
│   │    EmailService      │    │  TenantUserService   │        │
│   └──────────┬───────────┘    └──────────┬───────────┘        │
│              │                           │                     │
│              └───────────┬───────────────┘                     │
│                          │                                     │
│              ┌───────────▼───────────┐                         │
│              │   Shared Components   │                         │
│              │  - Rate limiter       │                         │
│              │  - Retry logic        │                         │
│              │  - Graph API wrapper  │                         │
│              └───────────────────────┘                         │
│                                                                │
└────────────────────────────────────────────────────────────────┘
```

## Alternatives considered

### Single unified service with mode flag

A single `CalendarService` that internally switches between delegated and app-only tokens
based on configuration or method parameters.

**Rejected because:**
- Mixes user-scoped and tenant-scoped semantics in one API
- Method signatures become confusing (when is `userId` the host's ID vs Microsoft UPN?)
- Harder to reason about which auth mode is active
- Testing requires mocking both paths

### Runtime mode switching

Allow the host to switch between modes at runtime or per-request.

**Rejected because:**
- Complicates token management (which cache applies?)
- Security risk: accidental tenant-wide access when user-scoped intended
- No clear use case justifies the complexity

### Separate module for app-only

A distinct `MicrosoftOutlookTenantModule` that handles only app-only auth.

**Rejected because:**
- Duplicates shared infrastructure (rate limiting, retries, Graph wrapper)
- Forces hosts using both modes to configure two modules
- Namespace collisions if both modules are imported

## Consequences

### Positive

- **Clear separation:** Developers know exactly which services to inject for their use case
- **No breaking changes:** Existing delegated auth users are unaffected
- **Type safety:** Compiler prevents mixing user-scoped and tenant-scoped operations
- **Testability:** Each mode can be tested in isolation
- **Security clarity:** App-only requires explicit opt-in and admin consent

### Negative

- **Surface area:** More services to document and maintain
- **Potential confusion:** New developers must understand when to use which service
- **Partial duplication:** Some business logic is similar between CalendarService and
  TenantCalendarService (mitigated by shared internal helpers)

### Operational

- Both modes share rate limiting infrastructure, so tenant-wide operations count against
  the same Graph throttling budgets
- App-only tokens are cached per-process; in multi-container deployments each container
  fetches its own token (acceptable given token TTL)

## Review trigger

Revisit if:
- Microsoft introduces a unified auth flow that obsoletes client credentials
- Significant logic duplication emerges between user-scoped and tenant-scoped services
- Customers request runtime mode switching with a compelling use case
