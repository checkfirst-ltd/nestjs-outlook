---
dep:
  type: reference
  audience: [app-developer, ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/services/subscription/microsoft-subscription.service.ts
  tags: [subscriptions, webhooks, service, api]
  links:
    - target: ../how-to/subscribe-to-webhooks.md
      rel: NEXT
    - target: ./event-types.md
      rel: USES
---

# MicrosoftSubscriptionService Reference

Injectable service that manages Microsoft Graph webhook subscriptions. Exported from
`@checkfirst/nestjs-outlook`. Graph subscriptions expire after ~3 days; the module renews
them on a schedule.

## Lifecycle methods

| Method | Signature | Returns | Notes |
|--------|-----------|---------|-------|
| `createWebhookSubscription` | `(externalUserId: string)` | `Promise<Subscription>` | Creates a delegated (`/me/events`) Graph subscription and stores it with a generated `clientState`. Removes any existing calendar subscription for the user first (see invariant below). |
| `createAppOnlyWebhookSubscription` | `(options: AppOnlySubscriptionOptions)` | `Promise<Subscription>` | Creates a tenant-wide app-only (`/users/{id}/events`) subscription. Removes any existing calendar subscription for the user first (see invariant below). |
| `renewWebhookSubscription` | `(subscriptionId: string, internalUserId: number)` | `Promise<Subscription>` | Renews an existing subscription; recreates it on `404`. |
| `deleteSubscription` | `(subscriptionId: string, accessToken: string)` | `Promise<void>` | Deletes a single subscription at Graph. |
| `deleteWebhookSubscription` | `(...)` | `Promise<void>` | Deletes a subscription at Graph and in the local DB. |
| `deleteAllWebhookSubscriptions` | `(userId: string \| number)` | `Promise<BulkSubscriptionDeleteResult>` | Removes all subscriptions for a user (Graph + DB). |
| `revokeTokens` | `(refreshToken: string)` | `Promise<void>` | Revokes the user's refresh token at Microsoft. |

### Invariant: one active calendar subscription per user

A user is identified across auth modes by the same internal `userId` on every subscription row
(delegated rows have `tenantId = null`; app-only rows carry `tenantId` + `microsoftUserId`). Both
`createWebhookSubscription` and `createAppOnlyWebhookSubscription` therefore remove any
already-active **calendar** subscription for that user (resource ending in `/events`) before
creating the new one — deleting it at Microsoft Graph with the token matching the old
subscription's own auth mode, and deactivating it locally.

This prevents duplicate notifications when a user connected via delegated OAuth is later mapped
into a tenant (or vice-versa). It is best-effort: a failed Graph delete still deactivates the row
locally (the subscription expires at Microsoft within ≤3 days) and never blocks creation of the
new subscription. Email subscriptions (`/me/messages`) are left untouched.

## Query methods

| Method | Signature | Returns |
|--------|-----------|---------|
| `getActiveSubscriptions` | `(accessToken: string)` | `Promise<MicrosoftSubscription[]>` |
| `getActiveSubscriptionsForClient` | `(...)` | `Promise<MicrosoftSubscription[]>` |
| `getActiveSubscriptionsForUser` | `(...)` | `Promise<MicrosoftSubscription[]>` |
| `getSubscription` | `(subscriptionId: string)` | `Promise<OutlookWebhookSubscription \| null>` |
| `getActiveSubscriptionForUser` | `(externalUserId: string)` | `Promise<string \| null>` |

## Cleanup methods

| Method | Signature | Returns |
|--------|-----------|---------|
| `cleanupSubscriptions` | `({ accessToken, filter? })` | `Promise<SubscriptionCleanupResult>` |
| `cleanupSubscriptionsForClient` | `(...)` | `Promise<SubscriptionCleanupResult>` |
| `cleanupSubscriptionsForUser` | `(userId: number, accessToken: string)` | `Promise<SubscriptionCleanupResult>` |
| `cleanupSubscriptionsForUserAndResource` | `(externalUserId: string, resource: string)` | `Promise<SubscriptionCleanupResult>` |

## Observability

| Method | Signature | Description |
|--------|-----------|-------------|
| `trackNotificationReceived` | `(subscriptionId: string): void` | Records that a notification arrived for a subscription. |

## Used by

- [Subscribe to webhooks](../how-to/subscribe-to-webhooks.md).
- [Event types](event-types.md) — events emitted on subscription lifecycle changes.
