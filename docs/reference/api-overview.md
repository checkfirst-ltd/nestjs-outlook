---
dep:
  type: reference
  audience: [ai-agent, app-developer]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/index.ts
    - src/microsoft-outlook.module.ts
  tags: [api, overview, exports, map, navigation]
  links:
    - target: ./configuration.md
      rel: USES
    - target: ./microsoft-auth-service.md
      rel: USES
    - target: ./calendar-service.md
      rel: USES
    - target: ./email-service.md
      rel: USES
    - target: ./subscription-service.md
      rel: USES
    - target: ./permission-scopes.md
      rel: USES
    - target: ./event-types.md
      rel: USES
    - target: ./http-endpoints.md
      rel: USES
    - target: ../explanation/architecture-overview.md
      rel: NEXT
---

# API Overview

Complete map of the public surface exported by `@checkfirst/nestjs-outlook` (see `src/index.ts`).
Each entry links to its detailed reference.

## Module

| Export | Kind | Reference |
|--------|------|-----------|
| `MicrosoftOutlookModule` | NestJS module (`forRoot` / `forRootAsync`) | [Configuration](configuration.md) |
| `MicrosoftOutlookConfig` | Config interface | [Configuration](configuration.md) |

## Services

| Service | Responsibility | Key methods | Reference |
|---------|----------------|-------------|-----------|
| `MicrosoftAuthService` | OAuth flow, token retrieval | `getLoginUrl`, `exchangeCodeForToken`, `getUserAccessToken` | [Auth](microsoft-auth-service.md) |
| `CalendarService` | Calendar events | `createEvent`, `updateEvent`, `getEventById`, `deleteEvent`, batch ops | [Calendar](calendar-service.md) |
| `EmailService` | Sending mail, mail webhooks | `sendEmail`, `createWebhookSubscription`, `handleEmailWebhook` | [Email](email-service.md) |
| `MicrosoftSubscriptionService` | Webhook subscription lifecycle | `createWebhookSubscription`, `renewWebhookSubscription`, `deleteAllWebhookSubscriptions`, cleanup | [Subscriptions](subscription-service.md) |

Additional exported services (advanced/internal): `RecurrenceService`, `UserIdConverterService`,
`DeltaSyncService`, `GraphRateLimiterService`, plus the `OutlookLockStore` and
`OutlookRateLimitStore` shared-state stores.

## Enums

| Enum | Purpose | Reference |
|------|---------|-----------|
| `PermissionScope` | Requestable permission scopes | [Permission scopes](permission-scopes.md) |
| `OutlookEventTypes` | Emitted event names | [Event types](event-types.md) |
| `ShowAsType` | Calendar free/busy status (mirrors Graph) | — |
| `MicrosoftUserStatus` | Stored user account state (`ACTIVE`, `CORRUPTED`, `SUBSCRIPTION_FAILED`) | — |

## HTTP controllers

| Controller | Base path | Reference |
|------------|-----------|-----------|
| `MicrosoftAuthController` | `auth/microsoft` | [HTTP endpoints](http-endpoints.md) |
| `CalendarController` | `calendar` | [HTTP endpoints](http-endpoints.md) |
| `EmailController` | `email` | [HTTP endpoints](http-endpoints.md) |

The `WebhookClientStateGuard` protects all webhook routes.

## Entities & repositories

| Export | Kind |
|--------|------|
| `MicrosoftUser` | Entity |
| `OutlookWebhookSubscription` | Entity |
| `OutlookWebhookSubscriptionRepository` | Repository |

## Errors

| Export | Thrown when |
|--------|-------------|
| `MailboxInactiveError` | The user's mailbox is not REST-enabled / inactive. |
| `CsrfValidationError` | OAuth `state` CSRF validation fails. |
| `SubscriptionSetupError` | Webhook subscription setup fails. |

## Constants & DI tokens

`MICROSOFT_CONFIG`, `OUTLOOK_LOCK_STORE`, `OUTLOOK_RATE_LIMIT_STORE`, `GRAPH_ERROR_CODES` —
see [Configuration](configuration.md).

## Deeper context

- [Architecture overview](../explanation/architecture-overview.md) — how these pieces fit together.
