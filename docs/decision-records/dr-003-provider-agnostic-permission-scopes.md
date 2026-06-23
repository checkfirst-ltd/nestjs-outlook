---
dep:
  type: decision-record
  audience: [library-contributor, ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: medium
  depends_on:
    - src/enums/permission-scope.enum.ts
    - src/services/auth/microsoft-auth.service.ts
  tags: [decision, permissions, scopes, abstraction]
  links:
    - target: ../reference/permission-scopes.md
      rel: DECIDES
    - target: ../reference/microsoft-auth-service.md
      rel: DECIDES
---

# DR-003: Provider-Agnostic Permission Scopes

## Context

Microsoft Graph scope strings (`Calendars.ReadWrite`, `Mail.Send`, `offline_access`, …) are
verbose, easy to mistype, and provider-specific. The host application needs to express *what
capability it wants* — read calendars, send mail — without embedding Microsoft's exact scope
vocabulary throughout its code.

## Decision

Expose a small `PermissionScope` enum of generic, capability-oriented values
(`CALENDAR_READ`, `CALENDAR_WRITE`, `EMAIL_READ`, `EMAIL_WRITE`, `EMAIL_SEND`) and map them
internally to the concrete Microsoft Graph scopes when building the authorization URL. The host
requests capabilities; the module owns the translation.

## Alternatives considered

- **Pass raw Graph scope strings through.** Rejected: leaks Microsoft's vocabulary into host
  code, invites typos, and makes the host responsible for required scopes like `offline_access`.
- **A single all-or-nothing scope set.** Rejected: violates least privilege — a calendar-only
  app should not request mail permissions.

## Consequences

- Host code is insulated from Graph scope naming and from mandatory scopes the module adds.
- The mapping is a single place to maintain; new capabilities require an enum addition plus a
  mapping entry.
- The abstraction is intentionally coarse — fine-grained Graph scopes not represented by the
  enum are not requestable without extending it.

## Review trigger

Revisit if the module must support additional Graph capabilities not covered by the enum, or if
multi-provider support is added (the generic enum was chosen partly to keep that door open).
