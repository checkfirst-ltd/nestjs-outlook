---
dep:
  type: reference
  audience:
    - library-integrator
    - app-developer
    - library-contributor
    - ai-agent
  owner: "@checkfirst-ltd"
  created: 2026-03-22
  last_verified: 2026-06-23T08:18:45.641Z
  confidence: high
  depends_on:
    - .docspec
  tags:
    - navigation
    - root
    - index
  links: []
---

# @checkfirst/nestjs-outlook — Documentation Root

> An opinionated NestJS module for Microsoft Outlook integration — OAuth authentication, calendar, email, and webhook subscriptions over the Microsoft Graph API. Supports both delegated (per-user) and app-only (tenant-wide) authentication.

---

## By Audience

### Library Integrator

**Entry point**: [getting-started.md](tutorials/getting-started.md)

- [Authenticate a User with Microsoft](how-to/authenticate-a-user.md)
- [Connect an Enterprise Tenant (App-Only Auth)](how-to/connect-enterprise-tenant.md)
- [Configure Redis Shared State](how-to/configure-redis-shared-state.md)
- [Getting Started: Connect Your First User to Microsoft](tutorials/getting-started.md)
- [HTTP Endpoints Reference](reference/http-endpoints.md)
- [Configuration Reference](reference/configuration.md)
- [PermissionScope Reference](reference/permission-scopes.md)

### Application Developer

**Entry point**: [index.md](index.md)

- [Handle Outlook Events](how-to/handle-outlook-events.md)
- [Authenticate a User with Microsoft](how-to/authenticate-a-user.md)
- [Connect an Enterprise Tenant (App-Only Auth)](how-to/connect-enterprise-tenant.md)
- [Configure Redis Shared State](how-to/configure-redis-shared-state.md)
- [Manage Calendar Events](how-to/manage-calendar-events.md)
- [Subscribe to Webhook Notifications](how-to/subscribe-to-webhooks.md)
- [Send an Email](how-to/send-email.md)
- [DR-002: Validate clientState on Webhook Endpoints](decision-records/dr-002-clientstate-webhook-validation.md)
- [HTTP Endpoints Reference](reference/http-endpoints.md)
- [API Overview](reference/api-overview.md)
- [RecurrenceService Reference](reference/recurrence-service.md)
- [EmailService Reference](reference/email-service.md)
- [ShowAsType Reference](reference/show-as-type.md)
- [MicrosoftUserStatus Reference](reference/microsoft-user-status.md)
- [UserIdConverterService Reference](reference/user-id-converter-service.md)
- [CalendarService Reference](reference/calendar-service.md)
- [TenantCalendarService Reference](reference/tenant-calendar-service.md)
- [TenantUserService Reference](reference/tenant-user-service.md)
- [AppOnlyAuthService Reference](reference/app-only-auth-service.md)
- [Tenant-Wide Authentication](explanation/tenant-wide-authentication.md)
- [Configuration Reference](reference/configuration.md)
- [MicrosoftSubscriptionService Reference](reference/subscription-service.md)
- [OutlookEventTypes Reference](reference/event-types.md)
- [MicrosoftAuthService Reference](reference/microsoft-auth-service.md)
- [PermissionScope Reference](reference/permission-scopes.md)

### Library Contributor

**Entry point**: [architecture-overview.md](explanation/architecture-overview.md)

- [Change Synchronization](explanation/change-synchronization.md)
- [Shared State and Concurrency](explanation/shared-state-and-concurrency.md)
- [Tenant-Wide Authentication](explanation/tenant-wide-authentication.md)
- [Architecture Overview](explanation/architecture-overview.md)
- [DR-001: Pluggable Shared-State Backend (Redis or In-Memory)](decision-records/dr-001-pluggable-shared-state-backend.md)
- [DR-003: Provider-Agnostic Permission Scopes](decision-records/dr-003-provider-agnostic-permission-scopes.md)
- [DR-002: Validate clientState on Webhook Endpoints](decision-records/dr-002-clientstate-webhook-validation.md)
- [DR-005: Per-User Throttling with a Service-Level Circuit Breaker](decision-records/dr-005-graph-throttling-circuit-breaker.md)
- [DR-004: Event-Driven Integration with the Host](decision-records/dr-004-event-driven-integration.md)
- [DR-006: Dual Authentication Architecture](decision-records/dr-006-dual-auth-architecture.md)
- [DR-007: Certificate Authentication](decision-records/dr-007-certificate-authentication.md)
- [DR-008: Tenant User Mapping](decision-records/dr-008-tenant-user-mapping.md)
- [Shared-State Stores Reference](reference/shared-state-stores.md)
- [RecurrenceService Reference](reference/recurrence-service.md)
- [MicrosoftUserStatus Reference](reference/microsoft-user-status.md)
- [UserIdConverterService Reference](reference/user-id-converter-service.md)
- [GraphRateLimiterService Reference](reference/graph-rate-limiter-service.md)
- [DeltaSyncService Reference](reference/delta-sync-service.md)
- [AppOnlyAuthService Reference](reference/app-only-auth-service.md)

### AI Agent

**Entry point**: [api-overview.md](reference/api-overview.md)

- [Change Synchronization](explanation/change-synchronization.md)
- [Shared State and Concurrency](explanation/shared-state-and-concurrency.md)
- [Tenant-Wide Authentication](explanation/tenant-wide-authentication.md)
- [Architecture Overview](explanation/architecture-overview.md)
- [DR-001: Pluggable Shared-State Backend (Redis or In-Memory)](decision-records/dr-001-pluggable-shared-state-backend.md)
- [DR-003: Provider-Agnostic Permission Scopes](decision-records/dr-003-provider-agnostic-permission-scopes.md)
- [DR-002: Validate clientState on Webhook Endpoints](decision-records/dr-002-clientstate-webhook-validation.md)
- [DR-005: Per-User Throttling with a Service-Level Circuit Breaker](decision-records/dr-005-graph-throttling-circuit-breaker.md)
- [DR-004: Event-Driven Integration with the Host](decision-records/dr-004-event-driven-integration.md)
- [DR-006: Dual Authentication Architecture](decision-records/dr-006-dual-auth-architecture.md)
- [DR-007: Certificate Authentication](decision-records/dr-007-certificate-authentication.md)
- [DR-008: Tenant User Mapping](decision-records/dr-008-tenant-user-mapping.md)
- [HTTP Endpoints Reference](reference/http-endpoints.md)
- [API Overview](reference/api-overview.md)
- [Shared-State Stores Reference](reference/shared-state-stores.md)
- [RecurrenceService Reference](reference/recurrence-service.md)
- [EmailService Reference](reference/email-service.md)
- [ShowAsType Reference](reference/show-as-type.md)
- [MicrosoftUserStatus Reference](reference/microsoft-user-status.md)
- [UserIdConverterService Reference](reference/user-id-converter-service.md)
- [CalendarService Reference](reference/calendar-service.md)
- [TenantCalendarService Reference](reference/tenant-calendar-service.md)
- [TenantUserService Reference](reference/tenant-user-service.md)
- [AppOnlyAuthService Reference](reference/app-only-auth-service.md)
- [Configuration Reference](reference/configuration.md)
- [MicrosoftSubscriptionService Reference](reference/subscription-service.md)
- [GraphRateLimiterService Reference](reference/graph-rate-limiter-service.md)
- [OutlookEventTypes Reference](reference/event-types.md)
- [MicrosoftAuthService Reference](reference/microsoft-auth-service.md)
- [DeltaSyncService Reference](reference/delta-sync-service.md)
- [PermissionScope Reference](reference/permission-scopes.md)

---

## By Type

### Explanation

- [Change Synchronization](explanation/change-synchronization.md)
- [Shared State and Concurrency](explanation/shared-state-and-concurrency.md)
- [Tenant-Wide Authentication](explanation/tenant-wide-authentication.md)
- [Architecture Overview](explanation/architecture-overview.md)

### How To

- [Handle Outlook Events](how-to/handle-outlook-events.md)
- [Authenticate a User with Microsoft](how-to/authenticate-a-user.md)
- [Connect an Enterprise Tenant (App-Only Auth)](how-to/connect-enterprise-tenant.md)
- [Configure Redis Shared State](how-to/configure-redis-shared-state.md)
- [Manage Calendar Events](how-to/manage-calendar-events.md)
- [Subscribe to Webhook Notifications](how-to/subscribe-to-webhooks.md)
- [Send an Email](how-to/send-email.md)

### Tutorial

- [Getting Started: Connect Your First User to Microsoft](tutorials/getting-started.md)

### Decision Record

- [DR-001: Pluggable Shared-State Backend (Redis or In-Memory)](decision-records/dr-001-pluggable-shared-state-backend.md)
- [DR-002: Validate clientState on Webhook Endpoints](decision-records/dr-002-clientstate-webhook-validation.md)
- [DR-003: Provider-Agnostic Permission Scopes](decision-records/dr-003-provider-agnostic-permission-scopes.md)
- [DR-004: Event-Driven Integration with the Host](decision-records/dr-004-event-driven-integration.md)
- [DR-005: Per-User Throttling with a Service-Level Circuit Breaker](decision-records/dr-005-graph-throttling-circuit-breaker.md)
- [DR-006: Dual Authentication Architecture](decision-records/dr-006-dual-auth-architecture.md)
- [DR-007: Certificate Authentication](decision-records/dr-007-certificate-authentication.md)
- [DR-008: Tenant User Mapping](decision-records/dr-008-tenant-user-mapping.md)

### Reference

- [HTTP Endpoints Reference](reference/http-endpoints.md)
- [API Overview](reference/api-overview.md)
- [Shared-State Stores Reference](reference/shared-state-stores.md)
- [RecurrenceService Reference](reference/recurrence-service.md)
- [EmailService Reference](reference/email-service.md)
- [ShowAsType Reference](reference/show-as-type.md)
- [MicrosoftUserStatus Reference](reference/microsoft-user-status.md)
- [MicrosoftTenantStatus Reference](reference/microsoft-tenant-status.md)
- [MicrosoftTenant Entity Reference](reference/microsoft-tenant-entity.md)
- [MicrosoftTenantUser Entity Reference](reference/microsoft-tenant-user-entity.md)
- [UserIdConverterService Reference](reference/user-id-converter-service.md)
- [CalendarService Reference](reference/calendar-service.md)
- [TenantCalendarService Reference](reference/tenant-calendar-service.md)
- [TenantUserService Reference](reference/tenant-user-service.md)
- [AppOnlyAuthService Reference](reference/app-only-auth-service.md)
- [Configuration Reference](reference/configuration.md)
- [MicrosoftSubscriptionService Reference](reference/subscription-service.md)
- [GraphRateLimiterService Reference](reference/graph-rate-limiter-service.md)
- [OutlookEventTypes Reference](reference/event-types.md)
- [MicrosoftAuthService Reference](reference/microsoft-auth-service.md)
- [DeltaSyncService Reference](reference/delta-sync-service.md)
- [PermissionScope Reference](reference/permission-scopes.md)
