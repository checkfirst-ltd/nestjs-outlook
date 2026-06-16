import { MockScenario, MockRequest, MockResponse } from '../interfaces';

let eventCounter = 0;
let subscriptionCounter = 0;

function mockEventId(): string {
  return `mock-event-${++eventCounter}-${Date.now()}`;
}

function mockSubscriptionId(): string {
  return `mock-sub-${++subscriptionCounter}-${Date.now()}`;
}

function buildEvent(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  const now = new Date().toISOString();
  const id = mockEventId();
  return {
    id,
    iCalUId: `ical-${id}`,
    subject: 'Mock Event',
    bodyPreview: 'This is a mock event',
    body: { contentType: 'html', content: '<p>Mock event body</p>' },
    start: { dateTime: now, timeZone: 'UTC' },
    end: { dateTime: new Date(Date.now() + 3600000).toISOString(), timeZone: 'UTC' },
    location: { displayName: 'Mock Location' },
    attendees: [],
    organizer: {
      emailAddress: { name: 'Mock User', address: 'mock@example.com' },
    },
    isAllDay: false,
    isCancelled: false,
    showAs: 'busy',
    sensitivity: 'normal',
    createdDateTime: now,
    lastModifiedDateTime: now,
    changeKey: `ck-${id}`,
    ...overrides,
  };
}

export const happyPathScenario: MockScenario = {
  name: 'happy-path',
  description: 'All API calls succeed with realistic sample data',
  routes: [
    // ── Token endpoints ──
    {
      method: 'POST',
      urlPattern: 'login.microsoftonline.com',
      handler: (): MockResponse => ({
        status: 200,
        data: {
          access_token: 'mock-access-token-' + Date.now(),
          refresh_token: 'mock-refresh-token-' + Date.now(),
          expires_in: 3600,
          token_type: 'Bearer',
          scope: 'Calendars.ReadWrite Mail.ReadWrite User.Read offline_access',
        },
      }),
    },

    // ── Mailbox validation ──
    {
      method: 'GET',
      urlPattern: '/me/mailboxSettings',
      handler: (): MockResponse => ({
        status: 200,
        data: {
          timeZone: 'UTC',
          dateFormat: 'MM/dd/yyyy',
          timeFormat: 'hh:mm tt',
          language: { locale: 'en-US' },
        },
      }),
    },

    // ── Default calendar ──
    {
      method: 'GET',
      urlPattern: /\/me\/calendar$/,
      handler: (): MockResponse => ({
        status: 200,
        data: {
          id: 'mock-default-calendar-id',
          name: 'Calendar',
          color: 'auto',
          isDefaultCalendar: true,
          canEdit: true,
          owner: { name: 'Mock User', address: 'mock@example.com' },
        },
      }),
    },

    // ── Calendar view (import / streaming) ──
    {
      method: 'GET',
      urlPattern: '/calendarView',
      handler: (): MockResponse => ({
        status: 200,
        data: {
          value: [buildEvent(), buildEvent({ subject: 'Mock Event 2' })],
          // No @odata.nextLink = last page
        },
      }),
    },

    // ── Delta sync ──
    {
      method: 'GET',
      urlPattern: '/events/delta',
      handler: (): MockResponse => ({
        status: 200,
        data: {
          value: [buildEvent({ subject: 'Delta Change Event' })],
          '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/me/events/delta?$deltatoken=mock-delta-token',
        },
      }),
    },

    // ── Recurring event instances ──
    {
      method: 'GET',
      urlPattern: '/instances',
      handler: (): MockResponse => ({
        status: 200,
        data: {
          value: [
            buildEvent({ subject: 'Instance 1', type: 'occurrence' }),
            buildEvent({ subject: 'Instance 2', type: 'occurrence' }),
          ],
        },
      }),
    },

    // ── Create event ──
    {
      method: 'POST',
      urlPattern: /\/calendars\/[^/]+\/events$/,
      handler: (req: MockRequest): MockResponse => ({
        status: 201,
        data: buildEvent(req.body as Record<string, unknown>),
      }),
    },

    // ── Update event ──
    {
      method: 'PATCH',
      urlPattern: /\/events\//,
      handler: (req: MockRequest): MockResponse => ({
        status: 200,
        data: buildEvent({
          ...(req.body as Record<string, unknown>),
          lastModifiedDateTime: new Date().toISOString(),
        }),
      }),
    },

    // ── Delete event ──
    {
      method: 'DELETE',
      urlPattern: /\/events\//,
      handler: (): MockResponse => ({
        status: 204,
        data: null,
      }),
    },

    // ── Get single event ──
    {
      method: 'GET',
      urlPattern: /\/events\/[^/?]+$/,
      handler: (): MockResponse => ({
        status: 200,
        data: buildEvent(),
      }),
    },

    // ── Batch operations ──
    {
      method: 'POST',
      urlPattern: '$batch',
      handler: (req: MockRequest): MockResponse => {
        const body = req.body as { requests: Array<{ id: string; method: string; url: string; body?: unknown }> };
        const responses = (body.requests || []).map((r) => {
          const isCreate = r.method === 'POST';
          const isDelete = r.method === 'DELETE';
          return {
            id: r.id,
            status: isDelete ? 204 : isCreate ? 201 : 200,
            body: isDelete ? null : buildEvent(r.body as Record<string, unknown> || {}),
          };
        });
        return { status: 200, data: { responses } };
      },
    },

    // ── Subscriptions ──
    {
      method: 'POST',
      urlPattern: '/subscriptions',
      handler: (req: MockRequest): MockResponse => {
        const body = req.body as Record<string, unknown>;
        const subId = mockSubscriptionId();
        return {
          status: 201,
          data: {
            id: subId,
            resource: body.resource || '/me/events',
            changeType: body.changeType || 'created,updated,deleted',
            clientState: body.clientState,
            notificationUrl: body.notificationUrl,
            expirationDateTime: body.expirationDateTime || new Date(Date.now() + 72 * 3600000).toISOString(),
          },
        };
      },
    },

    {
      method: 'GET',
      urlPattern: '/subscriptions',
      handler: (): MockResponse => ({
        status: 200,
        data: {
          value: [],
        },
      }),
    },

    {
      method: 'PATCH',
      urlPattern: /\/subscriptions\//,
      handler: (req: MockRequest): MockResponse => {
        const body = req.body as Record<string, unknown>;
        return {
          status: 200,
          data: {
            id: 'mock-sub-renewed',
            expirationDateTime: body.expirationDateTime || new Date(Date.now() + 72 * 3600000).toISOString(),
          },
        };
      },
    },

    {
      method: 'DELETE',
      urlPattern: /\/subscriptions\//,
      handler: (): MockResponse => ({
        status: 204,
        data: null,
      }),
    },

    // ── Send email ──
    {
      method: 'POST',
      urlPattern: '/me/sendMail',
      handler: (): MockResponse => ({
        status: 202,
        data: null,
      }),
    },

    // ── Get email message ──
    {
      method: 'GET',
      urlPattern: /\/me\/messages\//,
      handler: (): MockResponse => ({
        status: 200,
        data: {
          id: 'mock-message-id',
          subject: 'Mock Email',
          body: { contentType: 'html', content: '<p>Mock email body</p>' },
          from: { emailAddress: { name: 'Sender', address: 'sender@example.com' } },
          toRecipients: [{ emailAddress: { name: 'Recipient', address: 'recipient@example.com' } }],
          receivedDateTime: new Date().toISOString(),
          isRead: false,
        },
      }),
    },
  ],
};
