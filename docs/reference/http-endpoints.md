---
dep:
  type: reference
  audience: [library-integrator, app-developer, ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/controllers/microsoft-auth.controller.ts
    - src/controllers/calendar.controller.ts
    - src/controllers/email.controller.ts
    - src/guards/webhook-client-state.guard.ts
  tags: [http, controllers, endpoints, webhooks]
  links:
    - target: ../how-to/authenticate-a-user.md
      rel: NEXT
    - target: ../how-to/subscribe-to-webhooks.md
      rel: NEXT
---

# HTTP Endpoints Reference

Routes exposed by the module's bundled controllers. All paths are prefixed by the configured
`basePath` (e.g. `api/v1`).

## Authentication — `MicrosoftAuthController` (`auth/microsoft`)

| Method | Path | Query params | Description |
|--------|------|--------------|-------------|
| `GET` | `auth/microsoft/callback` | `code`, `state` | OAuth redirect target. Validates the CSRF token embedded in `state`, exchanges `code` for tokens, persists the user, and triggers subscription setup. |

## Calendar — `CalendarController` (`calendar`)

| Method | Path | Guard | Query params | Body | Description |
|--------|------|-------|--------------|------|-------------|
| `POST` | `calendar/webhook/notification` | `WebhookClientStateGuard` | `validationToken` | `OutlookWebhookNotificationDto` | Receives Graph change notifications for calendar resources. |
| `POST` | `calendar/webhook` | `WebhookClientStateGuard` | `validationToken` | `OutlookWebhookNotificationDto` | Alternate calendar webhook route. |

## Email — `EmailController` (`email`)

| Method | Path | Guard | Query params | Body | Description |
|--------|------|-------|--------------|------|-------------|
| `POST` | `email/webhook` | `WebhookClientStateGuard` | `validationToken` | `OutlookWebhookNotificationDto` | Receives Graph change notifications for mail resources. |

## Webhook validation handshake

When Microsoft Graph registers a subscription it sends a request carrying `validationToken`.
The controllers echo that token back as `text/plain` with HTTP `200` to confirm the endpoint.

## `WebhookClientStateGuard`

Applied to every webhook route. For each notification item it validates the `clientState`
against the stored subscription. Per-item rejection reasons:

| Reason | Meaning |
|--------|---------|
| `missing_subscription_id` | Notification item has no subscription ID. |
| `unknown_subscription` | No local subscription matches the ID. |
| `missing_stored_client_state` | The stored subscription has no `clientState`. |
| `client_state_mismatch` | The notification's `clientState` does not match the stored value. |

Rejected items emit the `WEBHOOK_REJECTED` event.

## Used by

- [Authenticate a user](../how-to/authenticate-a-user.md) — uses the callback route.
- [Subscribe to webhooks](../how-to/subscribe-to-webhooks.md) — uses the webhook routes.
