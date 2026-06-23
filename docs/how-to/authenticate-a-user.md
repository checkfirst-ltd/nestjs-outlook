---
dep:
  type: how-to
  audience: [app-developer, library-integrator]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/services/auth/microsoft-auth.service.ts
    - src/controllers/microsoft-auth.controller.ts
    - src/enums/permission-scope.enum.ts
  tags: [auth, oauth, login, tokens]
  links:
    - target: ../reference/microsoft-auth-service.md
      rel: USES
    - target: ../reference/permission-scopes.md
      rel: USES
    - target: ../reference/http-endpoints.md
      rel: USES
    - target: ../tutorials/getting-started.md
      rel: REQUIRES
---

# Authenticate a User with Microsoft

**Goal:** Send a user through the Microsoft OAuth flow and obtain an access token your code
can use to call Microsoft Graph on their behalf.

## Steps

### 1. Generate a login URL

Inject `MicrosoftAuthService` and call `getLoginUrl` with your application's identifier for
the user and the scopes you need.

```typescript
const loginUrl = await this.microsoftAuthService.getLoginUrl(externalUserId, [
  PermissionScope.CALENDAR_READ,
  PermissionScope.EMAIL_SEND,
]);
```

Redirect the user's browser to `loginUrl`.

### 2. Let the module handle the callback

The bundled `MicrosoftAuthController` exposes `GET {basePath}/auth/microsoft/callback`. When
Microsoft redirects the user back, the controller validates the CSRF token in the `state`
parameter and exchanges the `code` for tokens automatically. No code is required from you.

### 3. Get an access token for later calls

Whenever you need to call Graph for an authenticated user, request a fresh access token. The
service refreshes expired tokens transparently.

```typescript
const accessToken = await this.microsoftAuthService.getUserAccessToken({
  externalUserId,
});
```

## Verify

- After sign-in, confirm a row exists in `microsoft_users` for the user.
- Confirm `getUserAccessToken({ externalUserId })` returns a non-empty string.
- Listen for the `USER_AUTHENTICATED` event to confirm the flow completed.

## Related

- [MicrosoftAuthService reference](../reference/microsoft-auth-service.md) — full method signatures.
- [Permission scopes reference](../reference/permission-scopes.md) — available scope values.
