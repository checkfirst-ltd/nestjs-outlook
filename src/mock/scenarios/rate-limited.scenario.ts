import { MockScenario, MockResponse } from '../interfaces';

export const rateLimitedScenario: MockScenario = {
  name: 'rate-limited',
  description: 'All Graph API calls return 429 Too Many Requests with Retry-After header',
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
    // All Graph API calls get rate-limited
    {
      method: 'GET',
      urlPattern: 'graph.microsoft.com',
      handler: (): MockResponse => ({
        status: 429,
        data: {
          error: {
            code: 'TooManyRequests',
            message: 'Too many requests. Please retry after 10 seconds.',
          },
        },
        headers: { 'Retry-After': '10' },
      }),
    },
    {
      method: 'POST',
      urlPattern: 'graph.microsoft.com',
      handler: (): MockResponse => ({
        status: 429,
        data: {
          error: {
            code: 'TooManyRequests',
            message: 'Too many requests. Please retry after 10 seconds.',
          },
        },
        headers: { 'Retry-After': '10' },
      }),
    },
    {
      method: 'PATCH',
      urlPattern: 'graph.microsoft.com',
      handler: (): MockResponse => ({
        status: 429,
        data: {
          error: {
            code: 'TooManyRequests',
            message: 'Too many requests. Please retry after 10 seconds.',
          },
        },
        headers: { 'Retry-After': '10' },
      }),
    },
    {
      method: 'DELETE',
      urlPattern: 'graph.microsoft.com',
      handler: (): MockResponse => ({
        status: 429,
        data: {
          error: {
            code: 'TooManyRequests',
            message: 'Too many requests. Please retry after 10 seconds.',
          },
        },
        headers: { 'Retry-After': '10' },
      }),
    },
  ],
};
