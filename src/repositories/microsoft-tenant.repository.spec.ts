import { Test, TestingModule } from '@nestjs/testing';
import { getRepositoryToken } from '@nestjs/typeorm';
import { Repository } from 'typeorm';
import { MicrosoftTenantRepository } from './microsoft-tenant.repository';
import { MicrosoftTenant } from '../entities/microsoft-tenant.entity';
import { MicrosoftTenantStatus } from '../enums/microsoft-tenant-status.enum';

describe('MicrosoftTenantRepository', () => {
  let repository: MicrosoftTenantRepository;
  let mockTypeOrmRepository: jest.Mocked<Repository<MicrosoftTenant>>;

  const mockTenantId = '12345678-1234-1234-1234-123456789abc';
  const mockClientId = 'test-client-id';
  const mockThumbprint = 'abc123def456';

  const createMockTenant = (overrides?: Partial<MicrosoftTenant>): MicrosoftTenant => {
    const tenant = new MicrosoftTenant();
    tenant.id = 1;
    tenant.tenantId = mockTenantId;
    tenant.clientId = mockClientId;
    tenant.certificateThumbprint = mockThumbprint;
    tenant.certificatePath = '/path/to/cert.pem';
    tenant.certificateKeyPath = '/path/to/key.pem';
    tenant.status = MicrosoftTenantStatus.ACTIVE;
    tenant.isActive = true;
    tenant.adminConsentGrantedAt = new Date();
    tenant.createdAt = new Date();
    tenant.updatedAt = new Date();
    return Object.assign(tenant, overrides);
  };

  beforeEach(async () => {
    mockTypeOrmRepository = {
      findOne: jest.fn(),
      find: jest.fn(),
      save: jest.fn(),
      update: jest.fn(),
      delete: jest.fn(),
      count: jest.fn(),
      create: jest.fn(),
    } as unknown as jest.Mocked<Repository<MicrosoftTenant>>;

    const module: TestingModule = await Test.createTestingModule({
      providers: [
        MicrosoftTenantRepository,
        {
          provide: getRepositoryToken(MicrosoftTenant),
          useValue: mockTypeOrmRepository,
        },
      ],
    }).compile();

    repository = module.get<MicrosoftTenantRepository>(MicrosoftTenantRepository);
  });

  afterEach(() => {
    jest.clearAllMocks();
  });

  describe('save', () => {
    it('should save a new tenant connection', async () => {
      const newTenant: Partial<MicrosoftTenant> = {
        tenantId: mockTenantId,
        clientId: mockClientId,
        certificateThumbprint: mockThumbprint,
        status: MicrosoftTenantStatus.PENDING_CONSENT,
      };

      const savedTenant = createMockTenant(newTenant);

      mockTypeOrmRepository.findOne.mockResolvedValue(null);
      mockTypeOrmRepository.create.mockReturnValue(savedTenant);
      mockTypeOrmRepository.save.mockResolvedValue(savedTenant);

      const result = await repository.save(newTenant);

      expect(result).toEqual(savedTenant);
      expect(mockTypeOrmRepository.findOne).toHaveBeenCalledWith({
        where: { tenantId: mockTenantId },
      });
      expect(mockTypeOrmRepository.create).toHaveBeenCalled();
      expect(mockTypeOrmRepository.save).toHaveBeenCalled();
    });

    it('should update an existing tenant connection', async () => {
      const existingTenant = createMockTenant();
      const updateData: Partial<MicrosoftTenant> = {
        tenantId: mockTenantId,
        status: MicrosoftTenantStatus.ACTIVE,
      };

      const updatedTenant = createMockTenant({ status: MicrosoftTenantStatus.ACTIVE });

      mockTypeOrmRepository.findOne.mockResolvedValue(existingTenant);
      mockTypeOrmRepository.save.mockResolvedValue(updatedTenant);

      const result = await repository.save(updateData);

      expect(result).toEqual(updatedTenant);
      expect(mockTypeOrmRepository.findOne).toHaveBeenCalledWith({
        where: { tenantId: mockTenantId },
      });
      expect(mockTypeOrmRepository.save).toHaveBeenCalled();
      // Should preserve original ID
      expect(mockTypeOrmRepository.create).not.toHaveBeenCalled();
    });

    it('should create tenant without id field', async () => {
      const newTenantWithId: Partial<MicrosoftTenant> = {
        id: 999, // Should be stripped
        tenantId: mockTenantId,
        clientId: mockClientId,
      };

      const savedTenant = createMockTenant();

      mockTypeOrmRepository.findOne.mockResolvedValue(null);
      mockTypeOrmRepository.create.mockReturnValue(savedTenant);
      mockTypeOrmRepository.save.mockResolvedValue(savedTenant);

      await repository.save(newTenantWithId);

      // Verify create was called without id
      const createArg = mockTypeOrmRepository.create.mock.calls[0][0] as Partial<MicrosoftTenant>;
      expect(createArg.id).toBeUndefined();
    });
  });

  describe('findByTenantId', () => {
    it('should find tenant by Azure AD tenant ID', async () => {
      const tenant = createMockTenant();

      mockTypeOrmRepository.findOne.mockResolvedValue(tenant);

      const result = await repository.findByTenantId(mockTenantId);

      expect(result).toEqual(tenant);
      expect(mockTypeOrmRepository.findOne).toHaveBeenCalledWith({
        where: { tenantId: mockTenantId, isActive: true },
      });
    });

    it('should return null when tenant not found', async () => {
      mockTypeOrmRepository.findOne.mockResolvedValue(null);

      const result = await repository.findByTenantId('nonexistent-tenant');

      expect(result).toBeNull();
    });

    it('should cache tenant lookups', async () => {
      const tenant = createMockTenant();

      mockTypeOrmRepository.findOne.mockResolvedValue(tenant);

      // First call - should hit database
      const result1 = await repository.findByTenantId(mockTenantId);
      expect(result1).toEqual(tenant);
      expect(mockTypeOrmRepository.findOne).toHaveBeenCalledTimes(1);

      // Second call - should use cache
      const result2 = await repository.findByTenantId(mockTenantId);
      expect(result2).toEqual(tenant);
      expect(mockTypeOrmRepository.findOne).toHaveBeenCalledTimes(1); // Still 1
    });

    it('should not cache null results', async () => {
      mockTypeOrmRepository.findOne.mockResolvedValue(null);

      // First call
      await repository.findByTenantId(mockTenantId);
      expect(mockTypeOrmRepository.findOne).toHaveBeenCalledTimes(1);

      // Second call - should hit database again
      await repository.findByTenantId(mockTenantId);
      expect(mockTypeOrmRepository.findOne).toHaveBeenCalledTimes(2);
    });
  });

  describe('findByExternalTenantId', () => {
    it('should find tenant by external tenant ID (alias for findByTenantId)', async () => {
      const tenant = createMockTenant();

      mockTypeOrmRepository.findOne.mockResolvedValue(tenant);

      const result = await repository.findByExternalTenantId(mockTenantId);

      expect(result).toEqual(tenant);
      expect(mockTypeOrmRepository.findOne).toHaveBeenCalledWith({
        where: { tenantId: mockTenantId, isActive: true },
      });
    });
  });

  describe('updateStatus', () => {
    it('should update tenant status', async () => {
      mockTypeOrmRepository.update.mockResolvedValue({ affected: 1 } as any);

      await repository.updateStatus(mockTenantId, MicrosoftTenantStatus.CONSENT_REVOKED);

      expect(mockTypeOrmRepository.update).toHaveBeenCalledWith(
        { tenantId: mockTenantId },
        expect.objectContaining({
          status: MicrosoftTenantStatus.CONSENT_REVOKED,
          updatedAt: expect.any(Date),
        })
      );
    });

    it('should invalidate cache after status update', async () => {
      const tenant = createMockTenant();
      mockTypeOrmRepository.findOne.mockResolvedValue(tenant);
      mockTypeOrmRepository.update.mockResolvedValue({ affected: 1 } as any);

      // Populate cache
      await repository.findByTenantId(mockTenantId);
      expect(mockTypeOrmRepository.findOne).toHaveBeenCalledTimes(1);

      // Update status (should invalidate cache)
      await repository.updateStatus(mockTenantId, MicrosoftTenantStatus.DISABLED);

      // Next find should hit database again
      await repository.findByTenantId(mockTenantId);
      expect(mockTypeOrmRepository.findOne).toHaveBeenCalledTimes(2);
    });
  });

  describe('markConsentGranted', () => {
    it('should mark admin consent as granted', async () => {
      mockTypeOrmRepository.update.mockResolvedValue({ affected: 1 } as any);

      await repository.markConsentGranted(mockTenantId);

      expect(mockTypeOrmRepository.update).toHaveBeenCalledWith(
        { tenantId: mockTenantId },
        expect.objectContaining({
          status: MicrosoftTenantStatus.ACTIVE,
          adminConsentGrantedAt: expect.any(Date),
          updatedAt: expect.any(Date),
        })
      );
    });
  });

  describe('findAllActive', () => {
    it('should find all active tenants', async () => {
      const tenants = [
        createMockTenant({ id: 1, tenantId: 'tenant-1' }),
        createMockTenant({ id: 2, tenantId: 'tenant-2' }),
      ];

      mockTypeOrmRepository.find.mockResolvedValue(tenants);

      const result = await repository.findAllActive();

      expect(result).toEqual(tenants);
      expect(mockTypeOrmRepository.find).toHaveBeenCalledWith({
        where: {
          isActive: true,
          status: MicrosoftTenantStatus.ACTIVE,
        },
      });
    });
  });

  describe('deactivate', () => {
    it('should deactivate a tenant', async () => {
      mockTypeOrmRepository.update.mockResolvedValue({ affected: 1 } as any);

      await repository.deactivate(mockTenantId);

      expect(mockTypeOrmRepository.update).toHaveBeenCalledWith(
        { tenantId: mockTenantId },
        expect.objectContaining({
          isActive: false,
          updatedAt: expect.any(Date),
        })
      );
    });
  });

  describe('findByStatus', () => {
    it('should find tenants by status', async () => {
      const tenants = [createMockTenant({ status: MicrosoftTenantStatus.PENDING_CONSENT })];

      mockTypeOrmRepository.find.mockResolvedValue(tenants);

      const result = await repository.findByStatus(MicrosoftTenantStatus.PENDING_CONSENT);

      expect(result).toEqual(tenants);
      expect(mockTypeOrmRepository.find).toHaveBeenCalledWith({
        where: {
          status: MicrosoftTenantStatus.PENDING_CONSENT,
          isActive: true,
        },
      });
    });
  });

  describe('delete', () => {
    it('should delete a tenant permanently', async () => {
      const tenant = createMockTenant();

      mockTypeOrmRepository.findOne.mockResolvedValue(tenant);
      mockTypeOrmRepository.delete.mockResolvedValue({ affected: 1 } as any);

      await repository.delete(mockTenantId);

      expect(mockTypeOrmRepository.findOne).toHaveBeenCalledWith({
        where: { tenantId: mockTenantId },
      });
      expect(mockTypeOrmRepository.delete).toHaveBeenCalledWith({
        tenantId: mockTenantId,
      });
    });

    it('should not delete if tenant not found', async () => {
      mockTypeOrmRepository.findOne.mockResolvedValue(null);

      await repository.delete('nonexistent-tenant');

      expect(mockTypeOrmRepository.delete).not.toHaveBeenCalled();
    });

    it('should invalidate cache after delete', async () => {
      const tenant = createMockTenant();
      mockTypeOrmRepository.findOne.mockResolvedValue(tenant);
      mockTypeOrmRepository.delete.mockResolvedValue({ affected: 1 } as any);

      // Populate cache
      await repository.findByTenantId(mockTenantId);
      jest.clearAllMocks();
      mockTypeOrmRepository.findOne.mockResolvedValue(tenant);

      // Delete (should invalidate cache)
      await repository.delete(mockTenantId);

      // Next find should hit database again
      await repository.findByTenantId(mockTenantId);
      expect(mockTypeOrmRepository.findOne).toHaveBeenCalledTimes(2); // Once for delete check, once for find
    });
  });

  describe('countByStatus', () => {
    it('should count tenants by status', async () => {
      mockTypeOrmRepository.count.mockResolvedValue(5);

      const result = await repository.countByStatus(MicrosoftTenantStatus.ACTIVE);

      expect(result).toBe(5);
      expect(mockTypeOrmRepository.count).toHaveBeenCalledWith({
        where: {
          status: MicrosoftTenantStatus.ACTIVE,
          isActive: true,
        },
      });
    });
  });

  describe('cache behavior', () => {
    it('should invalidate cache when saving existing tenant', async () => {
      const existingTenant = createMockTenant();
      const updatedTenant = createMockTenant({ status: MicrosoftTenantStatus.DISABLED });

      // First, populate cache
      mockTypeOrmRepository.findOne.mockResolvedValue(existingTenant);
      await repository.findByTenantId(mockTenantId);
      expect(mockTypeOrmRepository.findOne).toHaveBeenCalledTimes(1);

      // Save (update) should invalidate cache
      mockTypeOrmRepository.findOne.mockResolvedValue(existingTenant);
      mockTypeOrmRepository.save.mockResolvedValue(updatedTenant);
      await repository.save({ tenantId: mockTenantId, status: MicrosoftTenantStatus.DISABLED });

      // Reset mock to return updated tenant
      mockTypeOrmRepository.findOne.mockResolvedValue(updatedTenant);

      // Next find should hit database again (cache was invalidated)
      const result = await repository.findByTenantId(mockTenantId);
      expect(result?.status).toBe(MicrosoftTenantStatus.DISABLED);
      // findOne called: 1 (initial) + 1 (save check) + 1 (after invalidation) = 3
      expect(mockTypeOrmRepository.findOne).toHaveBeenCalledTimes(3);
    });
  });
});
