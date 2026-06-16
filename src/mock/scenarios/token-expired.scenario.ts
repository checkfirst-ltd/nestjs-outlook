import { MockScenario, MockResponse } from '../interfaces';

export const tokenExpiredScenario: MockScenario = {
  name: 'token-expired',
  description: 'Token refresh fails with invalid_grant, simulating expired/revoked refresh token',
  routes: [
    {
      method: 'POST',
      urlPattern: 'login.microsoftonline.com',
      handler: (): MockResponse => ({
        status: 400,
        data: {
          error: 'invalid_grant',
          error_description:
            'AADSTS70000: The refresh token has expired due to inactivity.',
          error_codes: [70000],
          timestamp: new Date().toISOString(),
          trace_id: 'mock-trace-id',
          correlation_id: 'mock-correlation-id',
        },
      }),
    },
    // All Graph API calls return 401 Unauthorized
    {
      method: 'GET',
      urlPattern: 'graph.microsoft.com',
      handler: (): MockResponse => ({
        status: 401,
        data: {
          error: {
            code: 'InvalidAuthenticationToken',
            message: 'Access token has expired or is not yet valid.',
          },
        },
      }),
    },
    {
      method: 'POST',
      urlPattern: 'graph.microsoft.com',
      handler: (): MockResponse => ({
        status: 401,
        data: {
          error: {
            code: 'InvalidAuthenticationToken',
            message: 'Access token has expired or is not yet valid.',
          },
        },
      }),
    },
    {
      method: 'PATCH',
      urlPattern: 'graph.microsoft.com',
      handler: (): MockResponse => ({
        status: 401,
        data: {
          error: {
            code: 'InvalidAuthenticationToken',
            message: 'Access token has expired or is not yet valid.',
          },
        },
      }),
    },
    {
      method: 'DELETE',
      urlPattern: 'graph.microsoft.com',
      handler: (): MockResponse => ({
        status: 401,
        data: {
          error: {
            code: 'InvalidAuthenticationToken',
            message: 'Access token has expired or is not yet valid.',
          },
        },
      }),
    },
  ],
};
