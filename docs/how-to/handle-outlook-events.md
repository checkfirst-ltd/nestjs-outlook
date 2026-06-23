---
dep:
  type: how-to
  audience: [app-developer]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/enums/event-types.enum.ts
    - README.md
  tags: [events, event-emitter, notifications]
  links:
    - target: ../reference/event-types.md
      rel: USES
    - target: ../how-to/subscribe-to-webhooks.md
      rel: REQUIRES
---

# Handle Outlook Events

**Goal:** React in your application when the module reports a Microsoft activity — a new
calendar event, an incoming email, a completed authentication, or a subscription problem.

The module turns Graph webhook notifications and internal lifecycle changes into NestJS
events. You consume them with `@OnEvent` from `@nestjs/event-emitter`.

## Steps

### 1. Create a listener provider

```typescript
import { Injectable } from '@nestjs/common';
import { OnEvent } from '@nestjs/event-emitter';
import { OutlookEventTypes, OutlookResourceData } from '@checkfirst/nestjs-outlook';

@Injectable()
export class OutlookListener {
  @OnEvent(OutlookEventTypes.USER_AUTHENTICATED)
  onAuthenticated(externalUserId: string, data: { externalUserId: string; scopes: string[] }) {
    // A user finished connecting their Microsoft account.
  }

  @OnEvent(OutlookEventTypes.EVENT_CREATED)
  onEventCreated(data: OutlookResourceData) {
    // A calendar event was created for a connected user.
  }

  @OnEvent(OutlookEventTypes.EMAIL_RECEIVED)
  onEmailReceived(data: OutlookResourceData) {
    // A connected user received an email.
  }
}
```

### 2. Register the provider

Add `OutlookListener` to the `providers` array of one of your modules so Nest instantiates it
and binds the listeners.

### 3. React to revocation

Listen for `USER_REFRESH_TOKEN_INVALID` to prompt the user to reconnect when Microsoft
revokes access.

```typescript
@OnEvent(OutlookEventTypes.USER_REFRESH_TOKEN_INVALID)
onTokenInvalid(externalUserId: string) {
  // Mark the user as disconnected and ask them to sign in again.
}
```

## Verify

- Trigger an action in Outlook (create an event, send yourself an email) and confirm your
  handler runs.
- Confirm `USER_AUTHENTICATED` fires immediately after a user completes sign-in.

> Note: when an email is deleted, Microsoft sends a deletion notification followed by a
> creation one.

## Related

- [Event types reference](../reference/event-types.md) — every event name and when it fires.
- [Subscribe to webhooks](subscribe-to-webhooks.md) — the source of calendar and email events.
