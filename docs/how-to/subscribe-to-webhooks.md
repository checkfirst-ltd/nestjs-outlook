---
dep:
  type: how-to
  audience: [app-developer]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/services/subscription/microsoft-subscription.service.ts
    - src/controllers/calendar.controller.ts
    - src/controllers/email.controller.ts
    - src/guards/webhook-client-state.guard.ts
  tags: [webhooks, subscriptions, notifications, graph]
  links:
    - target: ../reference/subscription-service.md
      rel: USES
    - target: ../reference/http-endpoints.md
      rel: USES
    - target: ../how-to/handle-outlook-events.md
      rel: NEXT
    - target: ../how-to/authenticate-a-user.md
      rel: REQUIRES
---

# Subscribe to Webhook Notifications

**Goal:** Have Microsoft Graph push calendar and email change notifications to your app.

The module creates subscriptions **automatically** when a user authenticates, based on the
scopes you requested. This guide shows how to make that work and how to manage subscriptions
manually when you need to.

## Steps

### 1. Make your callback URL publicly reachable

Microsoft must be able to reach the webhook endpoints over HTTPS. Set `backendBaseUrl` in
`MicrosoftOutlookModule.forRoot()` to a public URL. For local development, expose your port
with a tunnel:

```bash
ngrok http 3000
```

Then set `backendBaseUrl` to the HTTPS URL ngrok prints.

### 2. Request scopes that imply subscriptions

When the user authenticates, request calendar and/or email scopes. On a successful token
exchange, the module creates the matching Graph subscriptions for that user. No extra call is
needed for the common case.

### 3. (Optional) Create a subscription manually

To (re)create a subscription outside the login flow, inject `MicrosoftSubscriptionService`.

```typescript
const subscription = await this.subscriptionService.createWebhookSubscription(externalUserId);
```

### 4. (Optional) Remove subscriptions

```typescript
// Remove every subscription for a user (Graph + local DB)
await this.subscriptionService.deleteAllWebhookSubscriptions(externalUserId);
```

Renewal is automatic: Graph subscriptions expire after ~3 days, and the module's scheduled
jobs renew or recreate them for active users.

## Verify

- A row appears in `outlook_webhook_subscriptions` for the user after authentication.
- Changing a calendar event or receiving an email triggers a `POST` to your webhook endpoint.
- Incoming notifications pass the `clientState` security check (rejected ones emit `WEBHOOK_REJECTED`).

## Related

- [Handle Outlook events](handle-outlook-events.md) — consume the events produced from notifications.
- [MicrosoftSubscriptionService reference](../reference/subscription-service.md) — all subscription methods.
- [HTTP endpoints reference](../reference/http-endpoints.md) — the webhook routes and validation handshake.
