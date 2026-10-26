---
dep:
  type: reference
  audience: [app-developer, ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/services/auth/microsoft-auth.service.ts
  tags: [auth, service, oauth, tokens, api]
  links:
    - target: ./permission-scopes.md
      rel: USES
    - target: ../how-to/authenticate-a-user.md
      rel: NEXT
---

# MicrosoftAuthService Reference

Injectable service that handles the Microsoft OAuth flow and access-token retrieval. Exported
from `@checkfirst/nestjs-outlook`.

## Methods

### `getLoginUrl(externalUserId, scopes?)`

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `externalUserId` | `string` | — | Host application's identifier for the user. |
| `scopes` | `PermissionScope[]` | configured default scopes | Generic permission scopes to request. |

**Returns:** `Promise<string>` — the Microsoft authorization URL to redirect the user to.
Embeds a CSRF token, the user ID, a timestamp, and the requested scopes in the `state` param.
The authority tenant comes from `delegatedAuth.tenant` and defaults to `common`; the same
authority is used for code exchange and token refresh. See the
[configuration reference](configuration.md#delegatedauthconfig).

### `exchangeCodeForToken(code, state)`

| Parameter | Type | Description |
|-----------|------|-------------|
| `code` | `string` | Authorization code from the OAuth callback. |
| `state` | `string` | The `state` value from the callback; CSRF and user ID are parsed from it. |

**Returns:** `Promise<TokenResponse>`. Throws if `state` is missing the user ID or fails CSRF
validation. Triggers subscription setup for the user based on the granted scopes.

### `getUserAccessToken(params)`

| Param field | Type | Default | Description |
|-------------|------|---------|-------------|
| `internalUserId` | `number` | — | Internal DB user ID (provide this or `externalUserId`). |
| `externalUserId` | `string` | — | Host application user ID. |
| `includeInactive` | `boolean` | `false` | Include users not in the `ACTIVE` state. |
| `cache` | `boolean` | `true` | Use the in-memory token cache. |

**Returns:** `Promise<string>` — a valid access token, refreshing it if expired. Throws
`MicrosoftRefreshTokenInvalidError` when the refresh token is no longer valid.

### `parseState(state)`

**Returns:** `StateObject | null`. Decodes the base64 `state` parameter into its object form.

### `validateCsrfToken(token, timestamp?)`

**Returns:** `Promise<string | null>` — an error string when invalid/expired, otherwise `null`.

### `cleanupExpiredTokens()`

Removes expired CSRF tokens from storage. Invoked on a schedule.

## Used by

- [Authenticate a user](../how-to/authenticate-a-user.md).
- [Permission scopes](permission-scopes.md) — the `scopes` argument values.
