import { EventEmitter2 } from '@nestjs/event-emitter';
import { TenantProvisioningService, BulkConnectUserInput } from './tenant-provisioning.service';
import { TenantUserService } from './tenant-user.service';
import { MicrosoftSubscriptionService } from '../subscription/microsoft-subscription.service';
import { MicrosoftTenantRepository } from '../../repositories/microsoft-tenant.repository';
import { MicrosoftTenant } from '../../entities/microsoft-tenant.entity';
import { MicrosoftUser } from '../../entities/microsoft-user.entity';
import { OutlookEventTypes } from '../../enums/event-types.enum';

describe('TenantProvisioningService', () => {
  let service: TenantProvisioningService;
  let tenantUserService: jest.Mocked<Pick<TenantUserService, 'registerUserMapping'>>;
  let subscriptionService: jest.Mocked<Pick<MicrosoftSubscriptionService, 'createAppOnlyWebhookSubscription'>>;
  let tenantRepository: jest.Mocked<Pick<MicrosoftTenantRepository, 'findByTenantId'>>;
  let eventEmitter: jest.Mocked<Pick<EventEmitter2, 'emit'>>;

  const tenantId = '12345678-1234-1234-1234-123456789abc';

  // registerUserMapping returns the mapped MicrosoftUser with a resolved microsoftUserId.
  const mappedUser = (externalUserId: string): MicrosoftUser =>
    ({ externalUserId, microsoftUserId: `ms-${externalUserId}` }) as MicrosoftUser;

  beforeEach(() => {
    tenantUserService = { registerUserMapping: jest.fn() };
    subscriptionService = { createAppOnlyWebhookSubscription: jest.fn() };
    tenantRepository = { findByTenantId: jest.fn().mockResolvedValue({ tenantId } as MicrosoftTenant) };
    eventEmitter = { emit: jest.fn() };

    service = new TenantProvisioningService(
      tenantUserService as unknown as TenantUserService,
      subscriptionService as unknown as MicrosoftSubscriptionService,
      tenantRepository as unknown as MicrosoftTenantRepository,
      eventEmitter as unknown as EventEmitter2,
    );
  });

  it('connects every user: maps + creates a subscription, and emits a completion event', async () => {
    const users: BulkConnectUserInput[] = [
      { externalUserId: 'insp-1', email: 'a@contoso.com' },
      { externalUserId: 'insp-2', email: 'b@contoso.com' },
      { externalUserId: 'insp-3', email: 'c@contoso.com' },
    ];
    tenantUserService.registerUserMapping.mockImplementation(async (_t, ext) => mappedUser(ext));
    subscriptionService.createAppOnlyWebhookSubscription.mockImplementation(async (opts) => ({
      id: `sub-${opts.externalUserId}`,
    }));

    const result = await service.connectUsers(tenantId, users);

    expect(result.total).toBe(3);
    expect(result.connected).toBe(3);
    expect(result.failed).toBe(0);
    // Each user was mapped then subscribed, with the resolved microsoftUserId.
    expect(tenantUserService.registerUserMapping).toHaveBeenCalledTimes(3);
    expect(tenantUserService.registerUserMapping).toHaveBeenCalledWith(tenantId, 'insp-1', 'a@contoso.com');
    expect(subscriptionService.createAppOnlyWebhookSubscription).toHaveBeenCalledWith({
      tenantId,
      microsoftUserId: 'ms-insp-1',
      externalUserId: 'insp-1',
    });
    // Results carry the created subscription id.
    expect(result.results.find((r) => r.externalUserId === 'insp-2')).toMatchObject({
      success: true,
      microsoftUserId: 'ms-insp-2',
      subscriptionId: 'sub-insp-2',
    });
    // Completion event emitted with the summary.
    expect(eventEmitter.emit).toHaveBeenCalledWith(
      OutlookEventTypes.TENANT_USERS_BULK_CONNECT_COMPLETED,
      expect.objectContaining({ tenantId, connected: 3, failed: 0 }),
    );
  });

  it('records a per-user failure without aborting the batch', async () => {
    const users: BulkConnectUserInput[] = [
      { externalUserId: 'ok-1', email: 'ok@contoso.com' },
      { externalUserId: 'missing', email: 'nope@contoso.com' },
    ];
    tenantUserService.registerUserMapping.mockImplementation(async (_t, ext, email) => {
      if (ext === 'missing') {
        throw new Error(`User not found in tenant: ${email}`);
      }
      return mappedUser(ext);
    });
    subscriptionService.createAppOnlyWebhookSubscription.mockResolvedValue({ id: 'sub-ok-1' });

    const result = await service.connectUsers(tenantId, users);

    expect(result.connected).toBe(1);
    expect(result.failed).toBe(1);
    const failed = result.results.find((r) => r.externalUserId === 'missing');
    expect(failed).toMatchObject({ success: false });
    expect(failed?.error).toContain('User not found in tenant');
    // The failing user never reached subscription creation; the other one did.
    expect(subscriptionService.createAppOnlyWebhookSubscription).toHaveBeenCalledTimes(1);
    // Completion event still emitted.
    expect(eventEmitter.emit).toHaveBeenCalledWith(
      OutlookEventTypes.TENANT_USERS_BULK_CONNECT_COMPLETED,
      expect.objectContaining({ connected: 1, failed: 1 }),
    );
  });

  it('fails fast and emits FAILED when the tenant is not found', async () => {
    tenantRepository.findByTenantId.mockResolvedValueOnce(null);

    await expect(
      service.connectUsers(tenantId, [{ externalUserId: 'x', email: 'x@contoso.com' }]),
    ).rejects.toThrow('Tenant not found');

    // No per-user work happened.
    expect(tenantUserService.registerUserMapping).not.toHaveBeenCalled();
    expect(subscriptionService.createAppOnlyWebhookSubscription).not.toHaveBeenCalled();
    expect(eventEmitter.emit).toHaveBeenCalledWith(
      OutlookEventTypes.TENANT_USERS_BULK_CONNECT_FAILED,
      expect.objectContaining({ tenantId }),
    );
  });

  it('marks a user failed when the mapping resolves no Microsoft user id', async () => {
    tenantUserService.registerUserMapping.mockResolvedValueOnce(
      { externalUserId: 'no-ms', microsoftUserId: null } as MicrosoftUser,
    );

    const result = await service.connectUsers(tenantId, [{ externalUserId: 'no-ms', email: 'x@contoso.com' }]);

    expect(result.failed).toBe(1);
    expect(subscriptionService.createAppOnlyWebhookSubscription).not.toHaveBeenCalled();
  });

  it('processes every user in a larger batch (bounded concurrency)', async () => {
    const users: BulkConnectUserInput[] = Array.from({ length: 23 }, (_, i) => ({
      externalUserId: `u-${i}`,
      email: `u${i}@contoso.com`,
    }));
    tenantUserService.registerUserMapping.mockImplementation(async (_t, ext) => mappedUser(ext));
    subscriptionService.createAppOnlyWebhookSubscription.mockResolvedValue({ id: 'sub' });

    const result = await service.connectUsers(tenantId, users);

    expect(result.total).toBe(23);
    expect(result.connected).toBe(23);
    expect(tenantUserService.registerUserMapping).toHaveBeenCalledTimes(23);
    expect(subscriptionService.createAppOnlyWebhookSubscription).toHaveBeenCalledTimes(23);
  });
});
