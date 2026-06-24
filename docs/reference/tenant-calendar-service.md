---
dep:
  type: reference
  audience: [app-developer, ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/services/calendar/tenant-calendar.service.ts
    - src/services/auth/app-only-auth.service.ts
  tags: [calendar, service, tenant, app-only, api]
  links:
    - target: ./app-only-auth-service.md
      rel: USES
    - target: ./calendar-service.md
      rel: RELATED
    - target: ../how-to/connect-enterprise-tenant.md
      rel: REQUIRES
---

# TenantCalendarService Reference

Injectable service for managing calendar events across all users in a Microsoft 365 tenant
using app-only authentication. Exported from `@checkfirst/nestjs-outlook`.

This service is only available when `appOnly.enabled` is `true` in the module configuration.
Unlike `CalendarService` which operates on behalf of a single authenticated user, this service
can access any user's calendar in the tenant.

## Methods

### `listEvents(userIdentifier, options?)`

Retrieves calendar events for any user in the tenant.

| Parameter | Type | Description |
|-----------|------|-------------|
| `userIdentifier` | `string` | User principal name (UPN) or user ID in Microsoft Graph |
| `options` | `ListEventsOptions` | Filter and pagination options |

**ListEventsOptions:**

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `startDateTime` | `Date` | — | Start of time range (required) |
| `endDateTime` | `Date` | — | End of time range (required) |
| `top` | `number` | `100` | Maximum events to return |
| `select` | `string[]` | all fields | Fields to include in response |
| `orderBy` | `string` | `'start/dateTime'` | Sort order |

**Returns:** `Promise<CalendarEvent[]>`

**Example:**
```typescript
const events = await this.tenantCalendar.listEvents('user@contoso.com', {
  startDateTime: new Date('2026-06-01'),
  endDateTime: new Date('2026-06-30'),
  top: 50,
});
```

### `getEvent(userIdentifier, eventId)`

Retrieves a specific calendar event.

| Parameter | Type | Description |
|-----------|------|-------------|
| `userIdentifier` | `string` | User principal name or user ID |
| `eventId` | `string` | Microsoft Graph event ID |

**Returns:** `Promise<CalendarEvent>`

### `createEvent(userIdentifier, event)`

Creates a calendar event for a user.

| Parameter | Type | Description |
|-----------|------|-------------|
| `userIdentifier` | `string` | User principal name or user ID |
| `event` | `CreateEventPayload` | Event details |

**CreateEventPayload:**

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `subject` | `string` | Yes | Event title |
| `start` | `DateTimeTimeZone` | Yes | Start date/time with timezone |
| `end` | `DateTimeTimeZone` | Yes | End date/time with timezone |
| `body` | `{ contentType: string; content: string }` | No | Event body/description |
| `location` | `{ displayName: string }` | No | Event location |
| `attendees` | `Attendee[]` | No | List of attendees |
| `isOnlineMeeting` | `boolean` | No | Create Teams meeting link |
| `showAs` | `ShowAsType` | No | Free/busy status |
| `recurrence` | `RecurrencePattern` | No | Recurrence settings |

**Returns:** `Promise<CalendarEvent>` — the created event with Graph-assigned ID.

**Example:**
```typescript
const event = await this.tenantCalendar.createEvent('user@contoso.com', {
  subject: 'Project Review',
  start: { dateTime: '2026-06-25T14:00:00', timeZone: 'UTC' },
  end: { dateTime: '2026-06-25T15:00:00', timeZone: 'UTC' },
  attendees: [
    { emailAddress: { address: 'colleague@contoso.com' }, type: 'required' }
  ],
  isOnlineMeeting: true,
});
```

### `updateEvent(userIdentifier, eventId, updates)`

Updates an existing calendar event.

| Parameter | Type | Description |
|-----------|------|-------------|
| `userIdentifier` | `string` | User principal name or user ID |
| `eventId` | `string` | Microsoft Graph event ID |
| `updates` | `Partial<CreateEventPayload>` | Fields to update |

**Returns:** `Promise<CalendarEvent>` — the updated event.

### `deleteEvent(userIdentifier, eventId)`

Deletes a calendar event.

| Parameter | Type | Description |
|-----------|------|-------------|
| `userIdentifier` | `string` | User principal name or user ID |
| `eventId` | `string` | Microsoft Graph event ID |

**Returns:** `Promise<void>`

### `getFreeBusy(userIdentifiers, timeRange)`

Retrieves free/busy information for multiple users.

| Parameter | Type | Description |
|-----------|------|-------------|
| `userIdentifiers` | `string[]` | Array of user principal names |
| `timeRange` | `{ start: Date; end: Date }` | Time range to check |

**Returns:** `Promise<FreeBusyResponse[]>` — availability for each user.

**Example:**
```typescript
const availability = await this.tenantCalendar.getFreeBusy(
  ['user1@contoso.com', 'user2@contoso.com'],
  { start: new Date(), end: new Date(Date.now() + 86400000) }
);

for (const user of availability) {
  console.log(`${user.userPrincipalName}: ${user.availabilityView}`);
}
```

### `listCalendars(userIdentifier)`

Lists all calendars for a user.

| Parameter | Type | Description |
|-----------|------|-------------|
| `userIdentifier` | `string` | User principal name or user ID |

**Returns:** `Promise<Calendar[]>` — list of user's calendars.

### `findMeetingTimes(userIdentifiers, options)`

Suggests meeting times based on attendee availability.

| Parameter | Type | Description |
|-----------|------|-------------|
| `userIdentifiers` | `string[]` | Attendee user principal names |
| `options` | `FindMeetingTimesOptions` | Constraints for meeting |

**Returns:** `Promise<MeetingTimeSuggestion[]>`

## User identifier formats

The `userIdentifier` parameter accepts:

- **User Principal Name (UPN):** `john.doe@contoso.com`
- **Microsoft Graph User ID:** `87d349ed-44d7-43e1-9a83-5f2406dee5bd`
- **Mail address:** `john.doe@contoso.com` (if different from UPN)

The service internally resolves these to the appropriate Graph API endpoint.

## Comparison with CalendarService

| Aspect | CalendarService | TenantCalendarService |
|--------|-----------------|----------------------|
| Auth type | Delegated (per-user OAuth) | App-only (client credentials) |
| Scope | Single authenticated user | Any user in tenant |
| Token source | User's refresh token | Application credentials |
| User identifier | `externalUserId` from your app | Microsoft UPN or user ID |
| Admin consent | Per-user consent | Tenant-wide admin consent |

## Error handling

| Error | Cause | Resolution |
|-------|-------|------------|
| `UserNotFoundError` | User identifier not found in tenant | Verify UPN or user ID |
| `CalendarNotFoundError` | User has no accessible calendar | Check mailbox provisioning |
| `InsufficientPermissionsError` | Missing `Calendars.ReadWrite` | Grant permission in Azure AD |
| `ThrottlingError` | Rate limit exceeded | Retry with backoff (handled internally) |

## Used by

- [Connect enterprise tenant](../how-to/connect-enterprise-tenant.md) — setup and usage guide

## Related

- [CalendarService reference](calendar-service.md) — delegated (per-user) calendar operations
- [AppOnlyAuthService reference](app-only-auth-service.md) — token acquisition
- [TenantUserService reference](tenant-user-service.md) — user enumeration
