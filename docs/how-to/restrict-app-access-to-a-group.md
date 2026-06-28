---
dep:
  type: how-to
  audience: [library-integrator, app-developer]
  owner: "@checkfirst-ltd"
  created: 2026-06-28
  last_verified: 2026-06-28T12:00:00+03:00
  confidence: high
  depends_on:
    - src/services/auth/app-only-auth.service.ts
    - src/services/tenant/tenant-user.service.ts
    - src/utils/retry.util.ts
  tags: [auth, enterprise, tenant, app-only, least-privilege, security-group, access-policy]
  links:
    - target: connect-enterprise-tenant.md
      rel: REQUIRES
    - target: ../reference/app-only-auth-service.md
      rel: USES
    - target: ../reference/tenant-calendar-service.md
      rel: USES
    - target: ../reference/permission-scopes.md
      rel: USES
    - target: ../decision-records/dr-006-dual-auth-architecture.md
      rel: EXPLAINS
---

# Restrict App Access to a Group

**Goal:** Let an enterprise customer's administrator confine this app's app-only (tenant-wide)
access to a specific set of users — the members of a security group they control — instead of
every mailbox in the tenant.

## Background

App-only authentication issues a token that, by default, can read and write **every** mailbox
and calendar in the consented tenant (see [Connect an Enterprise Tenant](connect-enterprise-tenant.md)).
That is rarely what a customer wants. Microsoft provides a tenant-side control so the
**customer's administrator** — not the app — decides which users the app may touch, managed
simply by adding or removing people from a group.

This scoping is enforced entirely on Microsoft's side. Nothing in this module's auth flow
changes: certificate → admin consent → token is identical. What changes is runtime behaviour —
calls for out-of-scope users return `403`, which the module already classifies as a permanent,
non-retryable error (see [retry.util.ts](../../src/utils/retry.util.ts)).

> **Who does what:** every step below is performed by the **enterprise customer's Exchange /
> Entra administrator** in *their* tenant. You (the app operator) only supply your app's
> **Application (client) ID**.

## Important scope boundary

The mechanisms below scope **Exchange data** — `Mail.ReadWrite` and `Calendars.ReadWrite`.
They do **not** restrict `User.Read.All`, which is a *directory* permission. With a mailbox
policy in place:

- ✅ Calendar / mail access is limited to the group's members.
- ⚠️ `User.Read.All` can still enumerate the whole directory (e.g. `TenantUserService.listUsers()`).

To keep the directory footprint minimal, prefer resolving users by the emails the customer has
explicitly mapped via `registerUserMapping` ([tenant-user.service.ts](../../src/services/tenant/tenant-user.service.ts))
rather than enumerating the tenant. See [Restrict directory reads](#restrict-directory-reads-optional).

## Option 1 — Application Access Policy (classic)

Widely supported and simple. The customer's Exchange admin runs the following in
**Exchange Online PowerShell**.

### 1. Create a mail-enabled security group and add the allowed users

```powershell
New-DistributionGroup `
  -Name "Outlook-App-Allowed" `
  -Type Security `
  -PrimarySmtpAddress outlook-app-allowed@customer.com

Add-DistributionGroupMember -Identity "Outlook-App-Allowed" -Member alice@customer.com
Add-DistributionGroupMember -Identity "Outlook-App-Allowed" -Member bob@customer.com
```

### 2. Restrict the app to that group

```powershell
New-ApplicationAccessPolicy `
  -AppId <your-app-client-id> `
  -PolicyScopeGroupId outlook-app-allowed@customer.com `
  -AccessRight RestrictAccess `
  -Description "Limit nestjs-outlook app to allowed users only"
```

`-AccessRight RestrictAccess` means: the app may access **only** members of the named group.
The admin manages access from then on purely by changing group membership — no further policy
edits, and nothing changes on the app side.

### 3. Verify the policy

```powershell
# Should report AccessCheckResult = Granted for an in-group user,
# and Denied for anyone outside the group.
Test-ApplicationAccessPolicy -Identity alice@customer.com  -AppId <your-app-client-id>
Test-ApplicationAccessPolicy -Identity carol@customer.com -AppId <your-app-client-id>
```

> **Propagation:** policy changes can take up to ~30 minutes to take effect across Exchange
> Online. Allow time before testing.

## Option 2 — RBAC for Applications (recommended direction)

Microsoft's newer, more scalable model (Application Access Policy has practical limits and is
being superseded). Configured in the **Exchange admin center** (*Roles → Admin roles*) or via
Graph: assign the app a role such as `Application Calendars.ReadWrite` **scoped to a security
group** (or an administrative unit). The end result is the same — the app's access is confined
to the group's members — but it scales better and is the path to prefer for new setups.

Use Option 1 if the customer's tooling/process already relies on it; otherwise prefer Option 2.

## Restrict directory reads (optional)

If the customer also wants to limit which users the app can *see* (not just which mailboxes it
can access), pick one of:

- Drop `User.Read.All` in favour of `User.ReadBasic.All`, or
- Scope directory access with **administrative units** + directory-scoped RBAC, or
- **(Recommended for this module)** have the host app look users up only by mapped emails via
  `registerUserMapping` / `getMicrosoftUserId` instead of calling `listUsers()`. The group
  policy on mail/calendar then becomes the effective boundary and a broad `User.Read.All` is
  never exercised tenant-wide.

## How the module behaves under scoping

- Token acquisition is unaffected — `getAccessToken()` still signs the assertion and returns a
  tenant-wide token ([app-only-auth.service.ts](../../src/services/auth/app-only-auth.service.ts)).
- Graph calls for **in-group** users succeed.
- Graph calls for **out-of-group** users return `403 ErrorAccessDenied`, treated as a permanent
  (non-retryable) error ([retry.util.ts](../../src/utils/retry.util.ts)). Handle this at the
  `TenantCalendarService` / email layer as an expected "user not in scope" outcome rather than a
  generic failure.

## Verify

- Run `Test-ApplicationAccessPolicy` (Option 1) for an in-group and an out-of-group user and
  confirm `Granted` / `Denied`.
- From the app, call `tenantCalendar.listEvents('alice@customer.com')` for an in-group user and
  confirm events return.
- Call the same for an out-of-group user and confirm a clean `403` is surfaced.

## Related

- [Connect an Enterprise Tenant](connect-enterprise-tenant.md) — the app-only setup this builds on
- [AppOnlyAuthService reference](../reference/app-only-auth-service.md) — token acquisition details
- [TenantCalendarService reference](../reference/tenant-calendar-service.md) — calendar operations
- [Permission scopes reference](../reference/permission-scopes.md) — required Graph permissions
- [DR-006: Dual Auth Architecture](../decision-records/dr-006-dual-auth-architecture.md) — design rationale