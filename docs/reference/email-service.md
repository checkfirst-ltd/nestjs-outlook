---
dep:
  type: reference
  audience: [app-developer, ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/services/email/email.service.ts
  tags: [email, service, api]
  links:
    - target: ../how-to/send-email.md
      rel: NEXT
---

# EmailService Reference

Injectable service for sending email and handling mail webhooks via Microsoft Graph. Exported
from `@checkfirst/nestjs-outlook`. The `Message` type is the Microsoft Graph `Message` model.

## Methods

### `sendEmail(message, externalUserId)`

| Parameter | Type | Description |
|-----------|------|-------------|
| `message` | `Partial<Message>` | Graph message payload (subject, body, recipients, …). |
| `externalUserId` | `string` | Host application user ID. |

**Returns:** `Promise<{ message: Message }>`. Sends via `/me/sendMail`; retries on transient
Graph errors (up to 7 attempts).

### `createWebhookSubscription(...)`

Creates a Graph subscription for the user's mail resource and stores it locally with a
generated `clientState`. **Returns:** the created subscription.

### `deleteWebhookSubscription(...)`

Deletes a mail subscription at Graph and in the local DB.

### `handleEmailWebhook(...)`

Processes an inbound mail notification: validates `clientState`, resolves the user, and emits
the matching email event (`EMAIL_RECEIVED`, `EMAIL_UPDATED`, or `EMAIL_DELETED`).

## Used by

- [Send an email](../how-to/send-email.md).
