---
dep:
  type: reference
  audience: [app-developer, ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/enums/event-types.enum.ts
  tags: [events, enum, event-emitter]
  links:
    - target: ../how-to/handle-outlook-events.md
      rel: NEXT
---

# OutlookEventTypes Reference

`OutlookEventTypes` is the enum of event names emitted on the NestJS event emitter. Subscribe
with `@OnEvent(OutlookEventTypes.X)`. Exported from `@checkfirst/nestjs-outlook`.

## Authentication events

| Member | Event name | Fires when |
|--------|-----------|------------|
| `USER_AUTHENTICATED` | `microsoft.auth.user.authenticated` | A user completes authentication. |
| `USER_REFRESH_TOKEN_INVALID` | `microsoft.auth.user.refresh_token_invalid` | A user's refresh token becomes invalid (revoked/expired). |

## Calendar events

| Member | Event name | Fires when |
|--------|-----------|------------|
| `EVENT_CREATED` | `outlook.event.created` | A calendar event is created. |
| `EVENT_UPDATED` | `outlook.event.updated` | A calendar event is updated. |
| `EVENT_DELETED` | `outlook.event.deleted` | A calendar event is deleted. |
| `EVENT_NOTIFICATION` | `outlook.event.notification` | A raw calendar notification is received. |
| `IMPORT_COMPLETED` | `outlook.calendar.import.completed` | A calendar import finishes. |

## Email events

| Member | Event name | Fires when |
|--------|-----------|------------|
| `EMAIL_RECEIVED` | `outlook.email.received` | A new email is received. |
| `EMAIL_UPDATED` | `outlook.email.updated` | An email is updated. |
| `EMAIL_DELETED` | `outlook.email.deleted` | An email is deleted. |

## Lifecycle events

| Member | Event name | Fires when |
|--------|-----------|------------|
| `LIFECYCLE_REAUTHORIZATION_REQUIRED` | `outlook.lifecycle.reauthorization_required` | Graph signals reauthorization is required. |
| `LIFECYCLE_SUBSCRIPTION_REMOVED` | `outlook.lifecycle.subscription_removed` | Graph reports a subscription was removed. |
| `LIFECYCLE_MISSED` | `outlook.lifecycle.missed` | A lifecycle notification was missed. |

## Subscription events

| Member | Event name | Fires when |
|--------|-----------|------------|
| `SUBSCRIPTION_RECREATED` | `outlook.subscription.recreated` | A subscription is recreated. |
| `SUBSCRIPTION_RECREATION_FAILED` | `outlook.subscription.recreation_failed` | Recreating a subscription fails. |
| `SUBSCRIPTION_AUTH_FAILED` | `outlook.subscription.auth_failed` | Subscription auth fails. |
| `SUBSCRIPTION_CREATION_FAILED` | `outlook.subscription.creation_failed` | Creating a subscription fails. |

## Cron observability events

| Member | Event name | Fires when |
|--------|-----------|------------|
| `HEALTH_CHECK_COMPLETED` | `outlook.cron.health_check.completed` | The health-check cron finishes. |
| `RETRY_FAILED_SUBSCRIPTIONS_COMPLETED` | `outlook.cron.retry_failed_subscriptions.completed` | The retry-failed-subscriptions cron finishes. |

## Security events

| Member | Event name | Fires when |
|--------|-----------|------------|
| `WEBHOOK_REJECTED` | `outlook.webhook.rejected` | A webhook notification fails the `clientState` check. |

## Used by

- [Handle Outlook events](../how-to/handle-outlook-events.md).
