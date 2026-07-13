import { EventEmitter2 } from '@nestjs/event-emitter';
import axios, { AxiosError } from 'axios';
import { MicrosoftSubscriptionService, AppOnlySubscriptionOptions } from './microsoft-subscription.service';
import { OutlookWebhookSubscriptionRepository } from '../../repositories/outlook-webhook-subscription.repository';
import { OutlookWebhookSubscription } from '../../entities/outlook-webhook-subscription.entity';
import { MicrosoftAuthService } from '../auth/microsoft-auth.service';
import { AppOnlyAuthService } from '../auth/app-only-auth.service';
import { UserIdConverterService } from '../shared/user-id-converter.service';
import { GraphRateLimiterService } from '../shared/graph-rate-limiter.service';
import { MicrosoftOutlookConfig } from '../../interfaces/config/outlook-config.interface';

jest.mock('axios');
const mockedAxios = axios as jest.Mocked<typeof axios>;

// Helper to create axios-like errors that pass axios.isAxiosError check
function createAxiosError(status: number, data?: unknown): AxiosError {
  const error = new Error() as AxiosError;
  error.isAxiosError = true;
  error.response = {
    status,
    data: data ?? {},
    statusText: '',
    headers: {},
    config: {} as any,
  };
  // Make axios.isAxiosError return true for this error
  (axios.isAxiosError as unknown as jest.Mock).mockImplementation(
    (err: unknown) => err === error || (err as any)?.isAxiosError === true
  );
  return error;
}

/**
 * Tests for app-only webhook subscription methods in MicrosoftSubscriptionService.
 *
 * These tests cover the tenant-wide subscription functionality that uses
 * app-only (client credentials) authentication instead of delegated user auth.
 */
describe('MicrosoftSubscriptionService - App-Only Methods', () => {
  let service: MicrosoftSubscriptionService;
  let mockWebhookRepo: jest.Mocked<OutlookWebhookSubscriptionRepository>;
  let mockAppOnlyAuthService: jest.Mocked<AppOnlyAuthService>;
  let mockUserIdConverter: jest.Mocked<UserIdConverterService>;
  let mockRateLimiter: jest.Mocked<GraphRateLimiterService>;
  let mockEventEmitter: jest.Mocked<EventEmitter2>;

  const mockConfig: MicrosoftOutlookConfig = {
    clientId: 'test-client-id',
    clientSecret: 'test-client-secret',
    redirectPath: '/auth/callback',
    backendBaseUrl: 'https://api.example.com',
    basePath: 'api/v1',
    calendarWebhookPath: '/calendar/webhook',
  };

  const testTenantId = '12345678-1234-1234-1234-123456789abc';
  const testMicrosoftUserId = 'user-guid-12345';
  const testExternalUserId = 'ext-user-1';
  const testInternalUserId = 42;
  const testSubscriptionId = 'sub-12345-abcde';
  const testAccessToken = 'app-only-access-token-xyz';

  beforeEach(() => {
    jest.clearAllMocks();

    // Create mocks
    mockWebhookRepo = {
      saveSubscription: jest.fn(),
      findBySubscriptionId: jest.fn(),
      updateSubscriptionExpiration: jest.fn(),
      deactivateSubscription: jest.fn(),
      deactivateAllByTenantId: jest.fn().mockResolvedValue(0),
      findAllActiveByTenantId: jest.fn(),
      findActiveByTenantAndMicrosoftUser: jest.fn(),
      findActiveSubscriptions: jest.fn(),
      findActiveByUserId: jest.fn(),
    } as unknown as jest.Mocked<OutlookWebhookSubscriptionRepository>;

    mockAppOnlyAuthService = {
      getAccessToken: jest.fn().mockResolvedValue(testAccessToken),
      isEnabled: jest.fn().mockReturnValue(true),
    } as unknown as jest.Mocked<AppOnlyAuthService>;

    mockUserIdConverter = {
      externalToInternal: jest.fn().mockResolvedValue(testInternalUserId),
      internalToExternal: jest.fn().mockResolvedValue(testExternalUserId),
    } as unknown as jest.Mocked<UserIdConverterService>;

    mockRateLimiter = {
      acquirePermit: jest.fn().mockResolvedValue(undefined),
    } as unknown as jest.Mocked<GraphRateLimiterService>;

    mockEventEmitter = {
      emit: jest.fn(),
    } as unknown as jest.Mocked<EventEmitter2>;

    // Create mock MicrosoftAuthService (not used in app-only methods but required)
    const mockMicrosoftAuthService = {} as unknown as MicrosoftAuthService;

    // Create mock MicrosoftUser repository
    const mockMicrosoftUserRepo = {
      findOne: jest.fn(),
    } as unknown as jest.Mocked<any>;

    // Create the service with mocks
    service = new MicrosoftSubscriptionService(
      mockMicrosoftAuthService,
      mockAppOnlyAuthService,
      mockWebhookRepo,
      mockEventEmitter,
      mockConfig,
      mockMicrosoftUserRepo,
      mockUserIdConverter,
      mockRateLimiter,
    );
  });

  describe('createAppOnlyWebhookSubscription', () => {
    const subscriptionOptions: AppOnlySubscriptionOptions = {
      tenantId: testTenantId,
      microsoftUserId: testMicrosoftUserId,
      externalUserId: testExternalUserId,
      internalUserId: testInternalUserId,
    };

    it('should create app-only webhook subscription with /users/{id}/events resource', async () => {
      const mockGraphResponse = {
        data: {
          id: testSubscriptionId,
          resource: `/users/${testMicrosoftUserId}/events`,
          changeType: 'created,updated,deleted',
          clientState: 'test-client-state',
          notificationUrl: 'https://api.example.com/api/v1/calendar/webhook',
          expirationDateTime: new Date(Date.now() + 72 * 60 * 60 * 1000).toISOString(),
        },
      };

      mockedAxios.post.mockResolvedValueOnce(mockGraphResponse);
      mockWebhookRepo.saveSubscription.mockResolvedValueOnce({
        id: 1,
        subscriptionId: testSubscriptionId,
        userId: testInternalUserId,
        tenantId: testTenantId,
        microsoftUserId: testMicrosoftUserId,
      } as OutlookWebhookSubscription);

      const result = await service.createAppOnlyWebhookSubscription(subscriptionOptions);

      // Verify app-only token was obtained
      expect(mockAppOnlyAuthService.getAccessToken).toHaveBeenCalledWith(testTenantId);

      // Verify Graph API call used correct resource path
      expect(mockedAxios.post).toHaveBeenCalledWith(
        'https://graph.microsoft.com/v1.0/subscriptions',
        expect.objectContaining({
          resource: `/users/${testMicrosoftUserId}/events`,
          changeType: 'created,updated,deleted',
        }),
        expect.objectContaining({
          headers: expect.objectContaining({
            Authorization: `Bearer ${testAccessToken}`,
            'Prefer': 'IdType="ImmutableId"',
          }),
        }),
      );

      // Verify subscription was saved with tenant context
      expect(mockWebhookRepo.saveSubscription).toHaveBeenCalledWith(
        expect.objectContaining({
          subscriptionId: testSubscriptionId,
          userId: testInternalUserId,
          tenantId: testTenantId,
          microsoftUserId: testMicrosoftUserId,
        }),
      );

      expect(result.id).toBe(testSubscriptionId);
      expect(result.resource).toBe(`/users/${testMicrosoftUserId}/events`);
    });

    it('should resolve internal user ID if not provided', async () => {
      const optionsWithoutInternalId: AppOnlySubscriptionOptions = {
        tenantId: testTenantId,
        microsoftUserId: testMicrosoftUserId,
        externalUserId: testExternalUserId,
      };

      mockedAxios.post.mockResolvedValueOnce({
        data: {
          id: testSubscriptionId,
          resource: `/users/${testMicrosoftUserId}/events`,
          clientState: 'test-state',
          expirationDateTime: new Date().toISOString(),
        },
      });
      mockWebhookRepo.saveSubscription.mockResolvedValueOnce({} as OutlookWebhookSubscription);

      await service.createAppOnlyWebhookSubscription(optionsWithoutInternalId);

      expect(mockUserIdConverter.externalToInternal).toHaveBeenCalledWith(
        testExternalUserId,
        { cache: false },
      );
    });

    it('should throw error when app-only auth is not configured', async () => {
      // Create service without AppOnlyAuthService
      const serviceWithoutAppOnly = new MicrosoftSubscriptionService(
        {} as MicrosoftAuthService,
        null, // No app-only auth service
        mockWebhookRepo,
        mockEventEmitter,
        mockConfig,
        {} as any,
        mockUserIdConverter,
        mockRateLimiter,
      );

      await expect(
        serviceWithoutAppOnly.createAppOnlyWebhookSubscription(subscriptionOptions),
      ).rejects.toThrow('App-only authentication is not configured');
    });

    it('should throw error when Graph API returns error', async () => {
      const axiosError = createAxiosError(403, {
        error: { code: 'AccessDenied', message: 'Insufficient privileges' },
      });
      mockedAxios.post.mockRejectedValueOnce(axiosError);

      await expect(
        service.createAppOnlyWebhookSubscription(subscriptionOptions),
      ).rejects.toThrow('Failed to create app-only webhook subscription');
    });
  });

  describe('renewAppOnlyWebhookSubscription', () => {
    it('should renew app-only subscription with tenant token', async () => {
      const newExpiration = new Date(Date.now() + 72 * 60 * 60 * 1000).toISOString();

      mockedAxios.patch.mockResolvedValueOnce({
        data: {
          id: testSubscriptionId,
          expirationDateTime: newExpiration,
        },
      });

      const result = await service.renewAppOnlyWebhookSubscription(
        testSubscriptionId,
        testTenantId,
      );

      // Verify app-only token was obtained
      expect(mockAppOnlyAuthService.getAccessToken).toHaveBeenCalledWith(testTenantId);

      // Verify Graph API call
      expect(mockedAxios.patch).toHaveBeenCalledWith(
        `https://graph.microsoft.com/v1.0/subscriptions/${testSubscriptionId}`,
        expect.objectContaining({
          expirationDateTime: expect.any(String),
        }),
        expect.objectContaining({
          headers: expect.objectContaining({
            Authorization: `Bearer ${testAccessToken}`,
          }),
        }),
      );

      // Verify local database was updated
      expect(mockWebhookRepo.updateSubscriptionExpiration).toHaveBeenCalledWith(
        testSubscriptionId,
        expect.any(Date),
      );

      expect(result.id).toBe(testSubscriptionId);
    });

    it('should deactivate subscription when Microsoft returns 404', async () => {
      const axiosError = createAxiosError(404);
      mockedAxios.patch.mockRejectedValueOnce(axiosError);

      await expect(
        service.renewAppOnlyWebhookSubscription(testSubscriptionId, testTenantId),
      ).rejects.toThrow(/not found at Microsoft/);

      expect(mockWebhookRepo.deactivateSubscription).toHaveBeenCalledWith(testSubscriptionId);
    });
  });

  describe('deleteAppOnlyWebhookSubscription', () => {
    it('should delete subscription at Microsoft and deactivate locally', async () => {
      mockedAxios.delete.mockResolvedValueOnce({ status: 204 });

      const result = await service.deleteAppOnlyWebhookSubscription(
        testSubscriptionId,
        testTenantId,
      );

      expect(mockAppOnlyAuthService.getAccessToken).toHaveBeenCalledWith(testTenantId);
      expect(mockedAxios.delete).toHaveBeenCalledWith(
        `https://graph.microsoft.com/v1.0/subscriptions/${testSubscriptionId}`,
        expect.objectContaining({
          headers: expect.objectContaining({
            Authorization: `Bearer ${testAccessToken}`,
          }),
        }),
      );
      expect(mockWebhookRepo.deactivateSubscription).toHaveBeenCalledWith(testSubscriptionId);
      expect(result).toBe(true);
    });

    it('should clean up locally even when Microsoft returns 404', async () => {
      const axiosError = createAxiosError(404);
      mockedAxios.delete.mockRejectedValueOnce(axiosError);

      const result = await service.deleteAppOnlyWebhookSubscription(
        testSubscriptionId,
        testTenantId,
      );

      expect(mockWebhookRepo.deactivateSubscription).toHaveBeenCalledWith(testSubscriptionId);
      expect(result).toBe(true);
    });
  });

  describe('deleteAllAppOnlySubscriptionsForTenant', () => {
    // Build a Graph $batch response body for the given per-request HTTP statuses.
    const batchResponse = (statuses: number[]) => ({
      data: {
        responses: statuses.map((status, i) => ({ id: `${i}`, status, body: null })),
      },
    });

    it('deletes all subscriptions via $batch and bulk-deactivates locally', async () => {
      const mockSubscriptions: Partial<OutlookWebhookSubscription>[] = [
        { subscriptionId: 'sub-1', tenantId: testTenantId, microsoftUserId: 'user-1' },
        { subscriptionId: 'sub-2', tenantId: testTenantId, microsoftUserId: 'user-2' },
        { subscriptionId: 'sub-3', tenantId: testTenantId, microsoftUserId: 'user-3' },
      ];

      mockWebhookRepo.findAllActiveByTenantId.mockResolvedValueOnce(
        mockSubscriptions as OutlookWebhookSubscription[],
      );
      mockedAxios.post.mockResolvedValueOnce(batchResponse([204, 204, 204]));
      mockWebhookRepo.deactivateAllByTenantId.mockResolvedValueOnce(3);

      const result = await service.deleteAllAppOnlySubscriptionsForTenant(testTenantId);

      expect(mockWebhookRepo.findAllActiveByTenantId).toHaveBeenCalledWith(testTenantId);
      expect(mockAppOnlyAuthService.getAccessToken).toHaveBeenCalledWith(testTenantId);

      // One $batch call, not one DELETE per subscription.
      expect(mockedAxios.post).toHaveBeenCalledTimes(1);
      expect(mockedAxios.post).toHaveBeenCalledWith(
        'https://graph.microsoft.com/v1.0/$batch',
        expect.objectContaining({
          requests: expect.arrayContaining([
            expect.objectContaining({ method: 'DELETE', url: '/subscriptions/sub-1' }),
          ]),
        }),
        expect.objectContaining({
          headers: expect.objectContaining({ Authorization: `Bearer ${testAccessToken}` }),
        }),
      );

      // Local rows deactivated in ONE statement, not per-subscription.
      expect(mockWebhookRepo.deactivateAllByTenantId).toHaveBeenCalledWith(testTenantId);
      expect(mockWebhookRepo.deactivateSubscription).not.toHaveBeenCalled();

      expect(result.totalFound).toBe(3);
      expect(result.successfullyDeleted).toBe(3);
      expect(result.failedToDelete).toBe(0);
      expect(result.deletedSubscriptionIds).toEqual(['sub-1', 'sub-2', 'sub-3']);
    });

    it('chunks into batches of 20 (Graph limit)', async () => {
      const mockSubscriptions = Array.from({ length: 45 }, (_, i) => ({
        subscriptionId: `sub-${i}`,
        tenantId: testTenantId,
      })) as OutlookWebhookSubscription[];

      mockWebhookRepo.findAllActiveByTenantId.mockResolvedValueOnce(mockSubscriptions);
      mockedAxios.post
        .mockResolvedValueOnce(batchResponse(new Array(20).fill(204)))
        .mockResolvedValueOnce(batchResponse(new Array(20).fill(204)))
        .mockResolvedValueOnce(batchResponse(new Array(5).fill(204)));
      mockWebhookRepo.deactivateAllByTenantId.mockResolvedValueOnce(45);

      const result = await service.deleteAllAppOnlySubscriptionsForTenant(testTenantId);

      // 45 subs → 3 batch calls (20 + 20 + 5), not 45 individual deletes.
      expect(mockedAxios.post).toHaveBeenCalledTimes(3);
      expect(result.successfullyDeleted).toBe(45);
      expect(mockWebhookRepo.deactivateAllByTenantId).toHaveBeenCalledWith(testTenantId);
    });

    it('should return empty result when no subscriptions exist', async () => {
      mockWebhookRepo.findAllActiveByTenantId.mockResolvedValueOnce([]);

      const result = await service.deleteAllAppOnlySubscriptionsForTenant(testTenantId);

      expect(result.totalFound).toBe(0);
      expect(result.successfullyDeleted).toBe(0);
      expect(result.deletedSubscriptionIds).toEqual([]);
      expect(mockedAxios.post).not.toHaveBeenCalled();
      expect(mockWebhookRepo.deactivateAllByTenantId).not.toHaveBeenCalled();
    });

    it('records per-item failures from the batch and still bulk-deactivates', async () => {
      const mockSubscriptions: Partial<OutlookWebhookSubscription>[] = [
        { subscriptionId: 'sub-1', tenantId: testTenantId },
        { subscriptionId: 'sub-2', tenantId: testTenantId },
      ];

      mockWebhookRepo.findAllActiveByTenantId.mockResolvedValueOnce(
        mockSubscriptions as OutlookWebhookSubscription[],
      );
      // sub-1 deleted (204), sub-2 fails (403).
      mockedAxios.post.mockResolvedValueOnce(batchResponse([204, 403]));
      mockWebhookRepo.deactivateAllByTenantId.mockResolvedValueOnce(2);

      const result = await service.deleteAllAppOnlySubscriptionsForTenant(testTenantId);

      expect(result.totalFound).toBe(2);
      expect(result.successfullyDeleted).toBe(1);
      expect(result.failedToDelete).toBe(1);
      expect(result.deletedSubscriptionIds).toEqual(['sub-1']);
      expect(result.errors[0].subscriptionId).toBe('sub-2');
      // Local rows still deactivated regardless of Microsoft outcome.
      expect(mockWebhookRepo.deactivateAllByTenantId).toHaveBeenCalledWith(testTenantId);
    });

    it('deactivates locally only when token acquisition fails (no batch call)', async () => {
      const mockSubscriptions: Partial<OutlookWebhookSubscription>[] = [
        { subscriptionId: 'sub-1', tenantId: testTenantId },
      ];

      mockWebhookRepo.findAllActiveByTenantId.mockResolvedValueOnce(
        mockSubscriptions as OutlookWebhookSubscription[],
      );
      mockAppOnlyAuthService.getAccessToken.mockRejectedValueOnce(
        new Error('Token acquisition failed'),
      );
      mockWebhookRepo.deactivateAllByTenantId.mockResolvedValueOnce(1);

      const result = await service.deleteAllAppOnlySubscriptionsForTenant(testTenantId);

      expect(mockedAxios.post).not.toHaveBeenCalled();
      expect(mockWebhookRepo.deactivateAllByTenantId).toHaveBeenCalledWith(testTenantId);
      expect(result.localOnlyDeactivated).toBe(1);
      expect(result.successfullyDeleted).toBe(0);
    });
  });

  describe('getActiveAppOnlySubscriptionForUser', () => {
    it('should get active subscription for specific user', async () => {
      const mockSubscription: Partial<OutlookWebhookSubscription> = {
        subscriptionId: testSubscriptionId,
        tenantId: testTenantId,
        microsoftUserId: testMicrosoftUserId,
        isActive: true,
      };

      mockWebhookRepo.findActiveByTenantAndMicrosoftUser.mockResolvedValueOnce(
        mockSubscription as OutlookWebhookSubscription,
      );

      const result = await service.getActiveAppOnlySubscriptionForUser(
        testTenantId,
        testMicrosoftUserId,
      );

      expect(mockWebhookRepo.findActiveByTenantAndMicrosoftUser).toHaveBeenCalledWith(
        testTenantId,
        testMicrosoftUserId,
      );
      expect(result).toBe(testSubscriptionId);
    });

    it('should return null when no active subscription exists', async () => {
      mockWebhookRepo.findActiveByTenantAndMicrosoftUser.mockResolvedValueOnce(null);

      const result = await service.getActiveAppOnlySubscriptionForUser(
        testTenantId,
        testMicrosoftUserId,
      );

      expect(result).toBeNull();
    });
  });

  describe('getAppOnlySubscriptionsForTenant', () => {
    it('should return all active subscriptions for tenant', async () => {
      const mockSubscriptions: Partial<OutlookWebhookSubscription>[] = [
        { subscriptionId: 'sub-1', tenantId: testTenantId, microsoftUserId: 'user-1' },
        { subscriptionId: 'sub-2', tenantId: testTenantId, microsoftUserId: 'user-2' },
      ];

      mockWebhookRepo.findAllActiveByTenantId.mockResolvedValueOnce(
        mockSubscriptions as OutlookWebhookSubscription[],
      );

      const result = await service.getAppOnlySubscriptionsForTenant(testTenantId);

      expect(mockWebhookRepo.findAllActiveByTenantId).toHaveBeenCalledWith(testTenantId);
      expect(result).toHaveLength(2);
    });
  });
});
