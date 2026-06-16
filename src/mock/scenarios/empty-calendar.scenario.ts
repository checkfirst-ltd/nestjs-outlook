import { MockScenario, MockResponse, MockRequest } from '../interfaces';

export const emptyCalendarScenario: MockScenario = {
  name: 'empty-calendar',
  description: 'User has a valid mailbox but zero events, no subscriptions, empty delta',
  routes: [
    // Token endpoint
    {
      method: 'POST',
      urlPattern: 'login.microsoftonline.com',
      handler: (): MockResponse => ({
        status: 200,
        data: {
          access_token: 'mock-access-token',
          refresh_token: 'mock-refresh-token',
          expires_in: 3600,
          token_type: 'Bearer',
        },
      }),
    },
    // Mailbox validation
    {
      method: 'GET',
      urlPattern: '/me/mailboxSettings',
      handler: (): MockResponse => ({
        status: 200,
        data: { timeZone: 'UTC' },
      }),
    },
    // Default calendar
    {
      method: 'GET',
      urlPattern: /\/me\/calendar$/,
      handler: (): MockResponse => ({
        status: 200,
        data: {
          id: 'mock-empty-calendar-id',
          name: 'Calendar',
          isDefaultCalendar: true,
        },
      }),
    },
    // Calendar view — empty
    {
      method: 'GET',
      urlPattern: '/calendarView',
      handler: (): MockResponse => ({
        status: 200,
        data: { value: [] },
      }),
    },
    // Delta sync — no changes
    {
      method: 'GET',
      urlPattern: '/events/delta',
      handler: (): MockResponse => ({
        status: 200,
        data: {
          value: [],
          '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/me/events/delta?$deltatoken=empty-token',
        },
      }),
    },
    // List subscriptions — none
    {
      method: 'GET',
      urlPattern: '/subscriptions',
      handler: (): MockResponse => ({
        status: 200,
        data: { value: [] },
      }),
    },
    // Create subscription — succeeds
    {
      method: 'POST',
      urlPattern: '/subscriptions',
      handler: (req: MockRequest): MockResponse => {
        const body = req.body as Record<string, unknown>;
        return {
          status: 201,
          data: {
            id: `mock-sub-empty-${Date.now()}`,
            resource: body.resource,
            changeType: body.changeType,
            clientState: body.clientState,
            notificationUrl: body.notificationUrl,
            expirationDateTime: body.expirationDateTime,
          },
        };
      },
    },
    // Delete subscription
    {
      method: 'DELETE',
      urlPattern: /\/subscriptions\//,
      handler: (): MockResponse => ({
        status: 204,
        data: null,
      }),
    },
  ],
};
