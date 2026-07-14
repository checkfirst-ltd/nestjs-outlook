import { EventEmitter2 } from '@nestjs/event-emitter';
import { HealthService } from './health.service';
import { TenantUserService } from '../tenant/tenant-user.service';
import { MicrosoftSubscriptionService } from '../subscription/microsoft-subscription.service';
import { OutlookWebhookSubscriptionRepository } from '../../repositories/outlook-webhook-subscription.repository';
import { MicrosoftUser } from '../../entities/microsoft-user.entity';
import { MicrosoftTenant } from '../../entities/microsoft-tenant.entity';
import { OutlookWebhookSubscription } from '../../entities/outlook-webhook-subscription.entity';
import { MicrosoftUserStatus } from '../../enums/microsoft-user-status.enum';
import { MicrosoftTenantStatus } from '../../enums/microsoft-tenant-status.enum';
import { UserHealthStatus } from '../../enums/user-health-status.enum';
import { OutlookEventTypes } from '../../enums/event-types.enum';

const HOUR = 60 * 60 * 1000;
const future = () => new Date(Date.now() + 72 * HOUR);
const past = (hours: number) => new Date(Date.now() - hours * HOUR);

describe('HealthService', () => {
  let service: HealthService;
  let tenantUserService: jest.Mocked<Pick<TenantUserService, 'findUsersByExternalIds'>>;
  let subscriptionService: jest.Mocked<
    Pick<
      MicrosoftSubscriptionService,
      'createWebhookSubscription' | 'createAppOnlyWebhookSubscription' | 'verifySubscriptionAtGraph'
    >
  >;
  let subscriptionRepo: jest.Mocked<Pick<OutlookWebhookSubscriptionRepository, 'findActiveByUserIds'>>;
  let eventEmitter: jest.Mocked<Pick<EventEmitter2, 'emit'>>;

  const tenantId = 'tenant-guid';

  // A delegated user (has a refresh token, no tenant mapping).
  const delegatedUser = (id: number, externalUserId: string, over: Partial<MicrosoftUser> = {}): MicrosoftUser =>
    ({
      id, externalUserId, isActive: true, status: MicrosoftUserStatus.ACTIVE,
      refreshToken: 'rt', microsoftUserId: null, tenant: null, ...over,
    }) as MicrosoftUser;

  // An app-only user (mapped into an ACTIVE tenant).
  const appOnlyUser = (id: number, externalUserId: string, over: Partial<MicrosoftUser> = {}): MicrosoftUser =>
    ({
      id, externalUserId, isActive: true, status: MicrosoftUserStatus.ACTIVE,
      refreshToken: null, microsoftUserId: `ms-${externalUserId}`,
      tenant: { tenantId, status: MicrosoftTenantStatus.ACTIVE, isActive: true } as MicrosoftTenant,
      ...over,
    }) as MicrosoftUser;

  const calendarSub = (userId: number, over: Partial<OutlookWebhookSubscription> = {}): OutlookWebhookSubscription =>
    ({
      subscriptionId: `sub-${userId}`, userId, resource: '/me/events', isActive: true,
      expirationDateTime: future(), lastNotificationAt: past(1), createdAt: past(1), tenantId: null, ...over,
    }) as OutlookWebhookSubscription;

  beforeEach(() => {
    tenantUserService = { findUsersByExternalIds: jest.fn().mockResolvedValue([]) };
    subscriptionService = {
      createWebhookSubscription: jest.fn().mockResolvedValue({ id: 'new-delegated' }),
      createAppOnlyWebhookSubscription: jest.fn().mockResolvedValue({ id: 'new-app-only' }),
      verifySubscriptionAtGraph: jest.fn().mockResolvedValue('present'),
    };
    subscriptionRepo = { findActiveByUserIds: jest.fn().mockResolvedValue([]) };
    eventEmitter = { emit: jest.fn() };

    service = new HealthService(
      tenantUserService as unknown as TenantUserService,
      subscriptionService as unknown as MicrosoftSubscriptionService,
      subscriptionRepo as unknown as OutlookWebhookSubscriptionRepository,
      eventEmitter as unknown as EventEmitter2,
    );
  });

  describe('checkUser — verdicts', () => {
    it('HEALTHY when active user has a live calendar subscription', async () => {
      tenantUserService.findUsersByExternalIds.mockResolvedValueOnce([delegatedUser(1, 'u1')]);
      subscriptionRepo.findActiveByUserIds.mockResolvedValueOnce([calendarSub(1)]);

      const health = await service.checkUser('u1');

      expect(health.status).toBe(UserHealthStatus.HEALTHY);
      expect(health.connected).toBe(true);
      expect(health.recoverable).toBe(false);
      expect(health.subscriptionId).toBe('sub-1');
    });

    it('NO_SUBSCRIPTION (recoverable) when the user has no active sub', async () => {
      tenantUserService.findUsersByExternalIds.mockResolvedValueOnce([appOnlyUser(2, 'u2')]);
      subscriptionRepo.findActiveByUserIds.mockResolvedValueOnce([]);

      const health = await service.checkUser('u2');

      expect(health.status).toBe(UserHealthStatus.NO_SUBSCRIPTION);
      expect(health.connected).toBe(false);
      expect(health.recoverable).toBe(true);
      expect(health.authMode).toBe('app-only');
    });

    it('SUBSCRIPTION_EXPIRED (recoverable) when the sub is past expiry', async () => {
      tenantUserService.findUsersByExternalIds.mockResolvedValueOnce([delegatedUser(1, 'u1')]);
      subscriptionRepo.findActiveByUserIds.mockResolvedValueOnce([calendarSub(1, { expirationDateTime: past(1) })]);

      const health = await service.checkUser('u1');

      expect(health.status).toBe(UserHealthStatus.SUBSCRIPTION_EXPIRED);
      expect(health.recoverable).toBe(true);
    });

    it('SUBSCRIPTION_STALE (recoverable) when no notification within the stale window', async () => {
      tenantUserService.findUsersByExternalIds.mockResolvedValueOnce([delegatedUser(1, 'u1')]);
      subscriptionRepo.findActiveByUserIds.mockResolvedValueOnce([
        calendarSub(1, { lastNotificationAt: past(48), createdAt: past(72) }),
      ]);

      const health = await service.checkUser('u1');

      expect(health.status).toBe(UserHealthStatus.SUBSCRIPTION_STALE);
      expect(health.recoverable).toBe(true);
    });

    it('NEEDS_REAUTH (not recoverable) when the delegated token is CORRUPTED', async () => {
      tenantUserService.findUsersByExternalIds.mockResolvedValueOnce([
        delegatedUser(1, 'u1', { status: MicrosoftUserStatus.CORRUPTED }),
      ]);
      subscriptionRepo.findActiveByUserIds.mockResolvedValueOnce([calendarSub(1)]);

      const health = await service.checkUser('u1');

      expect(health.status).toBe(UserHealthStatus.NEEDS_REAUTH);
      expect(health.recoverable).toBe(false);
    });

    it('NEEDS_ADMIN (not recoverable) when the app-only tenant is not ACTIVE', async () => {
      tenantUserService.findUsersByExternalIds.mockResolvedValueOnce([
        appOnlyUser(2, 'u2', {
          tenant: { tenantId, status: MicrosoftTenantStatus.CONSENT_REVOKED, isActive: false } as MicrosoftTenant,
        }),
      ]);

      const health = await service.checkUser('u2');

      expect(health.status).toBe(UserHealthStatus.NEEDS_ADMIN);
      expect(health.recoverable).toBe(false);
    });

    it('INACTIVE when the row is soft-deleted', async () => {
      tenantUserService.findUsersByExternalIds.mockResolvedValueOnce([delegatedUser(1, 'u1', { isActive: false })]);

      const health = await service.checkUser('u1');

      expect(health.status).toBe(UserHealthStatus.INACTIVE);
      expect(health.recoverable).toBe(false);
    });

    it('UNKNOWN when there is no user record', async () => {
      tenantUserService.findUsersByExternalIds.mockResolvedValueOnce([]);

      const health = await service.checkUser('ghost');

      expect(health.status).toBe(UserHealthStatus.UNKNOWN);
    });

    it('MISSING_AT_GRAPH when verifyAtGraph finds a DB-healthy sub gone at Microsoft', async () => {
      tenantUserService.findUsersByExternalIds.mockResolvedValueOnce([delegatedUser(1, 'u1')]);
      subscriptionRepo.findActiveByUserIds.mockResolvedValueOnce([calendarSub(1)]);
      subscriptionService.verifySubscriptionAtGraph.mockResolvedValueOnce('missing');

      const health = await service.checkUser('u1', { verifyAtGraph: true });

      expect(subscriptionService.verifySubscriptionAtGraph).toHaveBeenCalled();
      expect(health.status).toBe(UserHealthStatus.MISSING_AT_GRAPH);
      expect(health.recoverable).toBe(true);
    });

    it('stays HEALTHY when verifyAtGraph is inconclusive (unknown)', async () => {
      tenantUserService.findUsersByExternalIds.mockResolvedValueOnce([delegatedUser(1, 'u1')]);
      subscriptionRepo.findActiveByUserIds.mockResolvedValueOnce([calendarSub(1)]);
      subscriptionService.verifySubscriptionAtGraph.mockResolvedValueOnce('unknown');

      const health = await service.checkUser('u1', { verifyAtGraph: true });

      expect(health.status).toBe(UserHealthStatus.HEALTHY);
    });
  });

  describe('recoverUsers — auto-fix the fixable, report the rest', () => {
    it('routes recovery by auth mode and reports the unrecoverable', async () => {
      const users = [
        delegatedUser(1, 'healthy'),
        appOnlyUser(2, 'app-nosub'),
        delegatedUser(3, 'deleg-nosub'),
        delegatedUser(4, 'corrupted', { status: MicrosoftUserStatus.CORRUPTED }),
      ];
      tenantUserService.findUsersByExternalIds.mockResolvedValueOnce(users);
      // Only the healthy user has an active sub.
      subscriptionRepo.findActiveByUserIds.mockResolvedValueOnce([calendarSub(1)]);

      const result = await service.recoverUsers(['healthy', 'app-nosub', 'deleg-nosub', 'corrupted']);

      expect(result.total).toBe(4);
      expect(result.healthy).toBe(1);
      expect(result.recovered).toBe(2);
      expect(result.unrecoverable).toBe(1);
      expect(result.failed).toBe(0);

      // App-only recreate goes through the app-only path with the resolved ids.
      expect(subscriptionService.createAppOnlyWebhookSubscription).toHaveBeenCalledWith({
        tenantId,
        microsoftUserId: 'ms-app-nosub',
        externalUserId: 'app-nosub',
      });
      // Delegated recreate goes through the delegated path.
      expect(subscriptionService.createWebhookSubscription).toHaveBeenCalledWith('deleg-nosub');
      // CORRUPTED is reported, never re-created.
      expect(subscriptionService.createWebhookSubscription).not.toHaveBeenCalledWith('corrupted');
      expect(subscriptionService.createAppOnlyWebhookSubscription).toHaveBeenCalledTimes(1);
      expect(subscriptionService.createWebhookSubscription).toHaveBeenCalledTimes(1);

      // Completion event carries the summary.
      expect(eventEmitter.emit).toHaveBeenCalledWith(
        OutlookEventTypes.USER_HEALTH_RECOVERY_COMPLETED,
        expect.objectContaining({ recovered: 2, unrecoverable: 1 }),
      );
    });

    it('counts a recovery failure without aborting the batch', async () => {
      tenantUserService.findUsersByExternalIds.mockResolvedValueOnce([delegatedUser(3, 'deleg-nosub')]);
      subscriptionRepo.findActiveByUserIds.mockResolvedValueOnce([]);
      subscriptionService.createWebhookSubscription.mockRejectedValueOnce(new Error('graph down'));

      const result = await service.recoverUsers(['deleg-nosub']);

      expect(result.failed).toBe(1);
      expect(result.recovered).toBe(0);
      expect(result.results[0].action).toBe('recovery_failed');
      expect(result.results[0].recoveryError).toContain('graph down');
    });
  });
});
