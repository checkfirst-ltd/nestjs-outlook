---
dep:
  type: reference
  audience: [app-developer, ai-agent, library-integrator]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/enums/permission-scope.enum.ts
  tags: [permissions, scopes, enum, oauth]
  links:
    - target: ./microsoft-auth-service.md
      rel: USES
    - target: ../how-to/authenticate-a-user.md
      rel: NEXT
    - target: ../decision-records/dr-003-provider-agnostic-permission-scopes.md
      rel: NEXT
---

# PermissionScope Reference

`PermissionScope` is a provider-agnostic enum of permissions the host application can request.
Each value maps internally to the corresponding Microsoft Graph scope. Exported from
`@checkfirst/nestjs-outlook`.

## Values

| Member | String value | Grants |
|--------|--------------|--------|
| `CALENDAR_READ` | `CALENDAR_READ` | Read-only access to calendars. |
| `CALENDAR_WRITE` | `CALENDAR_WRITE` | Read-write access to calendars. |
| `EMAIL_READ` | `EMAIL_READ` | Read-only access to email. |
| `EMAIL_WRITE` | `EMAIL_WRITE` | Read-write access to email. |
| `EMAIL_SEND` | `EMAIL_SEND` | Permission to send email. |

## Notes

- Passed as the `scopes` argument to `MicrosoftAuthService.getLoginUrl`.
- Calendar scopes drive automatic calendar webhook subscription setup; email scopes drive
  automatic mail subscription setup at authentication time.

## Used by

- [Authenticate a user](../how-to/authenticate-a-user.md).
- [MicrosoftAuthService reference](microsoft-auth-service.md) — `getLoginUrl(scopes)`.
