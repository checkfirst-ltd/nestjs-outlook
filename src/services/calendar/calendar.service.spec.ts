import { CalendarService } from './calendar.service';

// Capture the authProvider handed to Client.init so we can invoke it like the Graph SDK would.
jest.mock('@microsoft/microsoft-graph-client', () => ({
  Client: {
    init: jest.fn((opts: { authProvider: unknown }) => ({ authProvider: opts.authProvider })),
  },
}));

type AuthProvider = (done: (error: Error | null, token: string | null) => void) => void;

/**
 * Regression tests for the per-request token refresh (createAuthRefreshingClient).
 *
 * The bug these guard against: freezing a single token at client construction. On long
 * reconciles the frozen token aged out before queued/retried requests were sent, and Microsoft
 * rejected it with "Lifetime validation failed, the token is expired". The auth provider must
 * therefore re-resolve the token on EVERY request, not capture it once.
 */
describe('CalendarService — per-request Graph auth', () => {
  let service: CalendarService;
  let mockMicrosoftAuthService: { getUserAccessToken: jest.Mock };
  let mockAppOnlyAuthService: { getAccessToken: jest.Mock };
  let mockMicrosoftUserRepo: { findOne: jest.Mock };

  const externalUserId = 'ext-user-1';
  const tenantId = 'tenant-guid-1';
  const appOnlyToken = 'app-only-token';
  const delegatedToken = 'delegated-token';

  // Drive the captured authProvider the way the Graph SDK does, and return the token it yields.
  const runAuthProvider = (client: { authProvider: AuthProvider }): Promise<string | null> =>
    new Promise((resolve, reject) => {
      client.authProvider((error, token) => {
        if (error) {
          reject(error);
        } else {
          resolve(token);
        }
      });
    });

  beforeEach(() => {
    jest.clearAllMocks();

    mockMicrosoftAuthService = {
      getUserAccessToken: jest.fn().mockResolvedValue(delegatedToken),
    };
    mockAppOnlyAuthService = {
      getAccessToken: jest.fn().mockResolvedValue(appOnlyToken),
    };
    mockMicrosoftUserRepo = { findOne: jest.fn() };

    service = new CalendarService(
      mockMicrosoftAuthService as never, // microsoftAuthService
      {} as never,                       // eventEmitter
      {} as never,                       // microsoftConfig
      {} as never,                       // deltaLinkRepository
      mockMicrosoftUserRepo as never,    // microsoftUserRepository
      {} as never,                       // deltaSyncService
      {} as never,                       // userIdConverter
      {} as never,                       // subscriptionService
      {} as never,                       // rateLimiter
      mockAppOnlyAuthService as never,   // appOnlyAuthService
    );
  });

  it('re-resolves the token on every request for a tenant (app-only) user', async () => {
    mockMicrosoftUserRepo.findOne.mockResolvedValue({
      microsoftUserId: 'ms-user-guid',
      tenant: { tenantId },
    });

    const client = (await service.getAuthenticatedClient(externalUserId)) as unknown as {
      authProvider: AuthProvider;
    };

    // Building the client must NOT pre-resolve a token — that is the frozen-token bug.
    expect(mockAppOnlyAuthService.getAccessToken).not.toHaveBeenCalled();

    // Each request re-resolves: two invocations → two fresh token acquisitions.
    await expect(runAuthProvider(client)).resolves.toBe(appOnlyToken);
    await expect(runAuthProvider(client)).resolves.toBe(appOnlyToken);

    expect(mockAppOnlyAuthService.getAccessToken).toHaveBeenCalledTimes(2);
    expect(mockAppOnlyAuthService.getAccessToken).toHaveBeenCalledWith({ tenantId });
    expect(mockMicrosoftAuthService.getUserAccessToken).not.toHaveBeenCalled();
  });

  it('uses the delegated token per request for a non-tenant user', async () => {
    mockMicrosoftUserRepo.findOne.mockResolvedValue({ microsoftUserId: null, tenant: null });

    const client = (await service.getAuthenticatedClient(externalUserId)) as unknown as {
      authProvider: AuthProvider;
    };

    await expect(runAuthProvider(client)).resolves.toBe(delegatedToken);
    await expect(runAuthProvider(client)).resolves.toBe(delegatedToken);

    expect(mockMicrosoftAuthService.getUserAccessToken).toHaveBeenCalledTimes(2);
    expect(mockAppOnlyAuthService.getAccessToken).not.toHaveBeenCalled();
  });

  it('propagates auth-resolution failure through the provider (so a request errors, not hangs)', async () => {
    mockMicrosoftUserRepo.findOne.mockResolvedValue({ microsoftUserId: null, tenant: null });
    mockMicrosoftAuthService.getUserAccessToken.mockRejectedValueOnce(new Error('token dead'));

    const client = (await service.getAuthenticatedClient(externalUserId)) as unknown as {
      authProvider: AuthProvider;
    };

    await expect(runAuthProvider(client)).rejects.toThrow('token dead');
  });
});
