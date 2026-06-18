import { MockScenario, MockResponse, MockRequest } from '../interfaces';

export const batchPartialFailureScenario: MockScenario = {
  name: 'batch-partial-failure',
  description: 'Batch operations return mixed results: odd-indexed requests fail with 409 Conflict',
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
        data: { id: 'mock-calendar-id', name: 'Calendar' },
      }),
    },
    // Batch operations — alternate success/failure
    {
      method: 'POST',
      urlPattern: '$batch',
      handler: (req: MockRequest): MockResponse => {
        const body = req.body as {
          requests: Array<{ id: string; method: string; url: string; body?: unknown }>;
        };
        const now = new Date().toISOString();
        const responses = (body.requests || []).map((r, index) => {
          if (index % 2 === 1) {
            // Odd-indexed items fail
            return {
              id: r.id,
              status: 409,
              body: {
                error: {
                  code: 'Conflict',
                  message: 'The resource has been modified by another process.',
                },
              },
            };
          }
          const isDelete = r.method === 'DELETE';
          const isCreate = r.method === 'POST';
          return {
            id: r.id,
            status: isDelete ? 204 : isCreate ? 201 : 200,
            body: isDelete
              ? null
              : {
                  id: `mock-event-batch-${r.id}`,
                  iCalUId: `ical-batch-${r.id}`,
                  subject: 'Batch Event',
                  createdDateTime: now,
                  lastModifiedDateTime: now,
                  ...(r.body as Record<string, unknown> || {}),
                },
          };
        });
        return { status: 200, data: { responses } };
      },
    },
    // Single event operations work normally
    {
      method: 'POST',
      urlPattern: /\/calendars\/[^/]+\/events$/,
      handler: (req: MockRequest): MockResponse => {
        const now = new Date().toISOString();
        return {
          status: 201,
          data: {
            id: `mock-event-${Date.now()}`,
            iCalUId: `ical-${Date.now()}`,
            createdDateTime: now,
            lastModifiedDateTime: now,
            ...(req.body as Record<string, unknown>),
          },
        };
      },
    },
    // Subscriptions work normally
    {
      method: 'GET',
      urlPattern: '/subscriptions',
      handler: (): MockResponse => ({
        status: 200,
        data: { value: [] },
      }),
    },
    {
      method: 'POST',
      urlPattern: '/subscriptions',
      handler: (): MockResponse => ({
        status: 201,
        data: {
          id: `mock-sub-${Date.now()}`,
          expirationDateTime: new Date(Date.now() + 72 * 3600000).toISOString(),
        },
      }),
    },
  ],
};
