import { EventEmitter2 } from '@nestjs/event-emitter';
import { MicrosoftSubscriptionService } from './microsoft-subscription.service';
import { OutlookWebhookSubscriptionRepository } from '../../repositories/outlook-webhook-subscription.repository';
import { MicrosoftAuthService } from '../auth/microsoft-auth.service';
import { AppOnlyAuthService } from '../auth/app-only-auth.service';
import { UserIdConverterService } from '../shared/user-id-converter.service';
import { GraphRateLimiterService } from '../shared/graph-rate-limiter.service';
import { MicrosoftOutlookConfig } from '../../interfaces/config/outlook-config.interface';

/**
 * Tests for deleteWebhookSubscription's failure policy.
 *
 * Disconnecting is a "make it so" operation: whatever Microsoft says, the caller has to end up
 * disconnected locally. These cover the paths that previously threw and left the subscription
 * active — a dead token, an unresolvable user, and a non-404 Graph error.
 */
describe('MicrosoftSubscriptionService.deleteWebhookSubscription', () => {
  let service: MicrosoftSubscriptionService;
  let mockWebhookRepo: jest.Mocked<OutlookWebhookSubscriptionRepository>;
  let mockMicrosoftAuthService: jest.Mocked<MicrosoftAuthService>;
  let mockUserIdConverter: jest.Mocked<UserIdConverterService>;
  let mockMicrosoftUserRepo: { findOne: jest.Mock; update: jest.Mock };
  let deleteAtGraph: jest.SpyInstance;

  const mockConfig: MicrosoftOutlookConfig = {
    clientId: 'test-client-id',
    clientSecret: 'test-client-secret',
    redirectPath: '/auth/callback',
    backendBaseUrl: 'https://api.example.com',
    basePath: 'api/v1',
    calendarWebhookPath: '/calendar/webhook',
  };

  const testSubscriptionId = 'sub-12345-abcde';
  const testExternalUserId = 'ext-user-1';
  const testInternalUserId = 42;
  const testAccessToken = 'delegated-user-token-abc';

  beforeEach(() => {
    jest.clearAllMocks();

    mockWebhookRepo = {
      deactivateSubscription: jest.fn().mockResolvedValue(undefined),
      count: jest.fn().mockResolvedValue(0),
    } as unknown as jest.Mocked<OutlookWebhookSubscriptionRepository>;

    mockUserIdConverter = {
      toInternalUserId: jest.fn().mockResolvedValue(testInternalUserId),
    } as unknown as jest.Mocked<UserIdConverterService>;

    mockMicrosoftAuthService = {
      getUserAccessToken: jest.fn().mockResolvedValue(testAccessToken),
    } as unknown as jest.Mocked<MicrosoftAuthService>;

    // No tenant relation → resolveUserAccessToken takes the delegated branch.
    mockMicrosoftUserRepo = {
      findOne: jest.fn().mockResolvedValue({ id: testInternalUserId, tenant: null }),
      update: jest.fn().mockResolvedValue({ affected: 1 }),
    };

    service = new MicrosoftSubscriptionService(
      mockMicrosoftAuthService,
      null as unknown as AppOnlyAuthService,
      mockWebhookRepo,
      { emit: jest.fn() } as unknown as EventEmitter2,
      mockConfig,
      mockMicrosoftUserRepo as never,
      mockUserIdConverter,
      { acquirePermit: jest.fn().mockResolvedValue(undefined) } as unknown as GraphRateLimiterService,
    );

    // Stub the raw Graph call — its retry/backoff behaviour is covered elsewhere and would
    // make these tests slow.
    deleteAtGraph = jest.spyOn(service, 'deleteSubscription').mockResolvedValue(undefined);
  });

  it('deletes at Microsoft and deactivates locally on the happy path', async () => {
    await expect(
      service.deleteWebhookSubscription(testSubscriptionId, testExternalUserId)
    ).resolves.toBe(true);

    expect(deleteAtGraph).toHaveBeenCalledWith(testSubscriptionId, testAccessToken);
    expect(mockWebhookRepo.deactivateSubscription).toHaveBeenCalledWith(testSubscriptionId);
    expect(mockMicrosoftUserRepo.update).toHaveBeenCalledWith(
      { externalUserId: testExternalUserId },
      { isActive: false }
    );
  });

  it('keeps the user active when another subscription remains', async () => {
    mockWebhookRepo.count.mockResolvedValueOnce(1);

    await expect(
      service.deleteWebhookSubscription(testSubscriptionId, testExternalUserId)
    ).resolves.toBe(true);

    expect(mockWebhookRepo.deactivateSubscription).toHaveBeenCalledWith(testSubscriptionId);
    expect(mockMicrosoftUserRepo.update).not.toHaveBeenCalled();
  });

  it('deactivates locally when the token cannot be refreshed', async () => {
    mockMicrosoftAuthService.getUserAccessToken.mockRejectedValueOnce(
      new Error('Failed to get valid access token: Failed to refresh access token from Microsoft')
    );

    await expect(
      service.deleteWebhookSubscription(testSubscriptionId, testExternalUserId)
    ).resolves.toBe(true);

    expect(deleteAtGraph).not.toHaveBeenCalled();
    expect(mockWebhookRepo.deactivateSubscription).toHaveBeenCalledWith(testSubscriptionId);
    expect(mockMicrosoftUserRepo.update).toHaveBeenCalled();
  });

  it('deactivates locally when the user has no delegated auth tokens', async () => {
    mockMicrosoftAuthService.getUserAccessToken.mockRejectedValueOnce(
      new Error(
        `Microsoft user ${testInternalUserId} has no delegated auth tokens (app-only user or not yet authenticated)`
      )
    );

    await expect(
      service.deleteWebhookSubscription(testSubscriptionId, testExternalUserId)
    ).resolves.toBe(true);

    expect(mockWebhookRepo.deactivateSubscription).toHaveBeenCalledWith(testSubscriptionId);
  });

  it('deactivates locally when Graph fails with a non-404 error', async () => {
    deleteAtGraph.mockRejectedValueOnce(new Error('HttpError 500'));

    await expect(
      service.deleteWebhookSubscription(testSubscriptionId, testExternalUserId)
    ).resolves.toBe(true);

    expect(mockWebhookRepo.deactivateSubscription).toHaveBeenCalledWith(testSubscriptionId);
    expect(mockMicrosoftUserRepo.update).toHaveBeenCalled();
  });

  it('deactivates the named subscription when the user cannot be resolved', async () => {
    mockUserIdConverter.toInternalUserId.mockRejectedValueOnce(
      new Error(`No active Microsoft user found for external ID: ${testExternalUserId}`)
    );

    await expect(
      service.deleteWebhookSubscription(testSubscriptionId, testExternalUserId)
    ).resolves.toBe(true);

    expect(mockWebhookRepo.deactivateSubscription).toHaveBeenCalledWith(testSubscriptionId);
    // Nothing else is safe to touch without a resolved user.
    expect(mockWebhookRepo.count).not.toHaveBeenCalled();
    expect(mockMicrosoftUserRepo.update).not.toHaveBeenCalled();
  });

  it('still throws when the local deactivation fails — the disconnect did not happen', async () => {
    const dbDown = new Error('ER_LOCK_WAIT_TIMEOUT');
    mockWebhookRepo.deactivateSubscription.mockRejectedValueOnce(dbDown);

    await expect(
      service.deleteWebhookSubscription(testSubscriptionId, testExternalUserId)
    ).rejects.toThrow(dbDown);
  });
});
