import { MockScenario, MockResponse } from '../interfaces';

export const mailboxInactiveScenario: MockScenario = {
  name: 'mailbox-inactive',
  description: 'Mailbox validation fails with MailboxNotEnabledForRESTAPI error (on-prem, soft-deleted, or inactive)',
  routes: [
    // Token endpoint works normally
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
    // Mailbox settings check fails
    {
      method: 'GET',
      urlPattern: '/me/mailboxSettings',
      handler: (): MockResponse => ({
        status: 403,
        data: {
          error: {
            code: 'MailboxNotEnabledForRESTAPI',
            message:
              "REST API is not yet supported for this mailbox. This could be because the user's mailbox is on a legacy Exchange server.",
            innerError: {
              date: new Date().toISOString(),
              'request-id': 'mock-request-id',
            },
          },
        },
      }),
    },
    // All other Graph calls also fail with same error
    {
      method: 'GET',
      urlPattern: 'graph.microsoft.com',
      handler: (): MockResponse => ({
        status: 403,
        data: {
          error: {
            code: 'MailboxNotEnabledForRESTAPI',
            message: 'REST API is not yet supported for this mailbox.',
          },
        },
      }),
    },
    {
      method: 'POST',
      urlPattern: 'graph.microsoft.com',
      handler: (): MockResponse => ({
        status: 403,
        data: {
          error: {
            code: 'MailboxNotEnabledForRESTAPI',
            message: 'REST API is not yet supported for this mailbox.',
          },
        },
      }),
    },
  ],
};
