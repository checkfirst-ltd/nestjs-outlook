---
dep:
  type: how-to
  audience: [app-developer]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/services/email/email.service.ts
  tags: [email, send, graph]
  links:
    - target: ../reference/email-service.md
      rel: USES
    - target: ../reference/permission-scopes.md
      rel: USES
    - target: ../how-to/authenticate-a-user.md
      rel: REQUIRES
---

# Send an Email

**Goal:** Send an email as an authenticated user through Microsoft Graph.

This step assumes the user connected their Microsoft account with the `EMAIL_SEND` scope.

## Steps

### 1. Build the message

Use the Microsoft Graph `Message` shape.

```typescript
const message = {
  subject: 'Hello from NestJS Outlook',
  body: {
    contentType: 'HTML',
    content: '<p>This is the email body</p>',
  },
  toRecipients: [
    { emailAddress: { address: 'recipient@example.com' } },
  ],
};
```

### 2. Send it

```typescript
const { message: sent } = await this.emailService.sendEmail(message, externalUserId);
```

`sendEmail` retries on transient Graph errors and respects the module's rate limiter, so you
do not need to add your own retry loop.

## Verify

- `sendEmail` resolves with a `message` object.
- The email appears in the user's **Sent Items** in Outlook.
- The recipient receives the message.

## Related

- [EmailService reference](../reference/email-service.md) — full method signature.
