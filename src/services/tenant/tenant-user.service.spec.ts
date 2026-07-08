import { Test, TestingModule } from '@nestjs/testing';
import { getRepositoryToken } from '@nestjs/typeorm';
import { Repository } from 'typeorm';
import axios from 'axios';
import { TenantUserService } from './tenant-user.service';
import { AppOnlyAuthService } from '../auth/app-only-auth.service';
import { MicrosoftTenant } from '../../entities/microsoft-tenant.entity';
import { MicrosoftUser } from '../../entities/microsoft-user.entity';
import { MICROSOFT_CONFIG } from '../../constants';
import { MicrosoftOutlookConfig } from '../../interfaces/config/outlook-config.interface';

jest.mock('axios');
const mockedAxios = axios as jest.Mocked<typeof axios>;

// executeGraphApiCall retries transient failures, so a plain Error would be retried
// (and, because jest.clearAllMocks keeps mock implementations, could hit a prior test's
// persistent success mock). Shape API failures as non-retryable Graph errors (status 403)
// so the service surfaces them immediately.
const graphError = (message: string, status = 403): Error =>
  Object.assign(new Error(message), { response: { status } });

describe('TenantUserService', () => {
  let service: TenantUserService;
  let appOnlyAuthService: jest.Mocked<AppOnlyAuthService>;
  let tenantRepository: jest.Mocked<Repository<MicrosoftTenant>>;
  let tenantUserRepository: jest.Mocked<Repository<MicrosoftUser>>;

  const mockTenantId = '12345678-1234-1234-1234-123456789abc';
  const mockMicrosoftUserId = 'user-guid-12345';
  const mockAccessToken = 'mock-access-token';
  const mockExternalUserId = 'app-user-123';
  const mockEmail = 'john.doe@contoso.com';
  const mockUpn = 'john.doe@contoso.com';

  const mockConfig: MicrosoftOutlookConfig = {
    clientId: 'test-client-id',
    clientSecret: 'test-client-secret',
    redirectPath: '/auth/callback',
    backendBaseUrl: 'https://api.example.com',
    appOnly: {
      enabled: true,
      tenantId: mockTenantId,
    },
  };

  const mockGraphUser = {
    id: mockMicrosoftUserId,
    userPrincipalName: mockUpn,
    displayName: 'John Doe',
    mail: mockEmail,
  };

  const mockTenant: Partial<MicrosoftTenant> = {
    id: 1,
    tenantId: mockTenantId,
    isActive: true,
  };

  const mockTenantUser: Partial<MicrosoftUser> = {
    id: 1,
    externalUserId: mockExternalUserId,
    microsoftUserId: mockMicrosoftUserId,
    userPrincipalName: mockUpn,
    isActive: true,
    tenant: mockTenant as MicrosoftTenant,
  };

  beforeEach(async () => {
    const mockAppOnlyAuthService = {
      getAccessToken: jest.fn().mockResolvedValue(mockAccessToken),
      isEnabled: jest.fn().mockReturnValue(true),
    };

    const mockTenantRepository = {
      findOne: jest.fn(),
    };

    const mockTenantUserRepository = {
      findOne: jest.fn(),
      save: jest.fn(),
    };

    const module: TestingModule = await Test.createTestingModule({
      providers: [
        TenantUserService,
        {
          provide: AppOnlyAuthService,
          useValue: mockAppOnlyAuthService,
        },
        {
          provide: getRepositoryToken(MicrosoftTenant),
          useValue: mockTenantRepository,
        },
        {
          provide: getRepositoryToken(MicrosoftUser),
          useValue: mockTenantUserRepository,
        },
        {
          provide: MICROSOFT_CONFIG,
          useValue: mockConfig,
        },
      ],
    }).compile();

    service = module.get<TenantUserService>(TenantUserService);
    appOnlyAuthService = module.get(AppOnlyAuthService);
    tenantRepository = module.get(getRepositoryToken(MicrosoftTenant));
    tenantUserRepository = module.get(getRepositoryToken(MicrosoftUser));

    jest.clearAllMocks();
    // Clear the internal cache before each test
    service.clearCache();
  });

  describe('lookupUserByEmail', () => {
    it('should lookup user by email address', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { value: [mockGraphUser] },
      });

      const result = await service.lookupUserByEmail(mockTenantId, mockEmail);

      expect(result).toEqual({
        microsoftUserId: mockMicrosoftUserId,
        userPrincipalName: mockUpn,
        displayName: 'John Doe',
        email: mockEmail,
      });
      expect(appOnlyAuthService.getAccessToken).toHaveBeenCalledWith(mockTenantId);
    });

    it('should include Prefer: IdType="ImmutableId" header', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { value: [mockGraphUser] },
      });

      await service.lookupUserByEmail(mockTenantId, mockEmail);

      expect(mockedAxios.get).toHaveBeenCalledWith(
        expect.any(String),
        expect.objectContaining({
          headers: expect.objectContaining({
            'Prefer': 'IdType="ImmutableId"',
          }),
        })
      );
    });

    it('should use $filter to search by mail or userPrincipalName', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { value: [mockGraphUser] },
      });

      await service.lookupUserByEmail(mockTenantId, mockEmail);

      expect(mockedAxios.get).toHaveBeenCalledWith(
        'https://graph.microsoft.com/v1.0/users',
        expect.objectContaining({
          params: expect.objectContaining({
            '$filter': expect.stringContaining('mail eq'),
          }),
        })
      );
    });

    it('should return null when user not found', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { value: [] },
      });

      const result = await service.lookupUserByEmail(mockTenantId, 'unknown@contoso.com');

      expect(result).toBeNull();
    });

    it('should cache lookup results', async () => {
      mockedAxios.get.mockResolvedValue({
        data: { value: [mockGraphUser] },
      });

      // First call
      await service.lookupUserByEmail(mockTenantId, mockEmail);

      // Second call should use cache
      await service.lookupUserByEmail(mockTenantId, mockEmail);

      expect(mockedAxios.get).toHaveBeenCalledTimes(1);
    });

    it('should throw error on Graph API failure', async () => {
      mockedAxios.get.mockRejectedValueOnce(graphError('Network error'));

      await expect(
        service.lookupUserByEmail(mockTenantId, mockEmail)
      ).rejects.toThrow('Failed to lookup user by email');
    });
  });

  describe('lookupUserByUpn', () => {
    it('should lookup user by UPN using direct endpoint', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: mockGraphUser,
      });

      const result = await service.lookupUserByUpn(mockTenantId, mockUpn);

      expect(result).toEqual({
        microsoftUserId: mockMicrosoftUserId,
        userPrincipalName: mockUpn,
        displayName: 'John Doe',
        email: mockEmail,
      });
    });

    it('should use /users/{upn} endpoint for direct lookup', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: mockGraphUser,
      });

      await service.lookupUserByUpn(mockTenantId, mockUpn);

      expect(mockedAxios.get).toHaveBeenCalledWith(
        expect.stringContaining(`/users/${encodeURIComponent(mockUpn)}`),
        expect.any(Object)
      );
    });

    it('should return null when user not found', async () => {
      mockedAxios.get.mockResolvedValueOnce(null);

      const result = await service.lookupUserByUpn(mockTenantId, 'unknown@contoso.com');

      expect(result).toBeNull();
    });

    it('should cache by both UPN and email', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: mockGraphUser,
      });

      // Lookup by UPN
      await service.lookupUserByUpn(mockTenantId, mockUpn);

      // Lookup by email should hit cache
      const result = await service.lookupUserByEmail(mockTenantId, mockEmail);

      expect(mockedAxios.get).toHaveBeenCalledTimes(1);
      expect(result?.microsoftUserId).toBe(mockMicrosoftUserId);
    });
  });

  describe('getUserById', () => {
    it('should get user by Microsoft Graph ID', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: mockGraphUser,
      });

      const result = await service.getUserById(mockTenantId, mockMicrosoftUserId);

      expect(result).toEqual({
        microsoftUserId: mockMicrosoftUserId,
        userPrincipalName: mockUpn,
        displayName: 'John Doe',
        email: mockEmail,
      });
    });

    it('should use /users/{id} endpoint', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: mockGraphUser,
      });

      await service.getUserById(mockTenantId, mockMicrosoftUserId);

      expect(mockedAxios.get).toHaveBeenCalledWith(
        `https://graph.microsoft.com/v1.0/users/${mockMicrosoftUserId}`,
        expect.any(Object)
      );
    });

    it('should return null for non-existent user', async () => {
      mockedAxios.get.mockResolvedValueOnce(null);

      const result = await service.getUserById(mockTenantId, 'non-existent-id');

      expect(result).toBeNull();
    });

    it('should select required fields only', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: mockGraphUser,
      });

      await service.getUserById(mockTenantId, mockMicrosoftUserId);

      expect(mockedAxios.get).toHaveBeenCalledWith(
        expect.any(String),
        expect.objectContaining({
          params: expect.objectContaining({
            '$select': 'id,userPrincipalName,displayName,mail',
          }),
        })
      );
    });
  });

  describe('registerUserMapping', () => {
    it('should create user mapping between external ID and Microsoft user', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { value: [mockGraphUser] },
      });
      tenantRepository.findOne.mockResolvedValueOnce(mockTenant as MicrosoftTenant);
      tenantUserRepository.findOne.mockResolvedValueOnce(null);
      tenantUserRepository.save.mockResolvedValueOnce(mockTenantUser as MicrosoftUser);

      const result = await service.registerUserMapping(
        mockTenantId,
        mockExternalUserId,
        mockEmail
      );

      expect(result.externalUserId).toBe(mockExternalUserId);
      expect(result.microsoftUserId).toBe(mockMicrosoftUserId);
      expect(tenantUserRepository.save).toHaveBeenCalled();
    });

    it('should lookup user before creating mapping', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { value: [mockGraphUser] },
      });
      tenantRepository.findOne.mockResolvedValueOnce(mockTenant as MicrosoftTenant);
      tenantUserRepository.findOne.mockResolvedValueOnce(null);
      tenantUserRepository.save.mockResolvedValueOnce(mockTenantUser as MicrosoftUser);

      await service.registerUserMapping(mockTenantId, mockExternalUserId, mockEmail);

      expect(mockedAxios.get).toHaveBeenCalled();
    });

    it('should throw error when user not found in tenant', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { value: [] },
      });

      await expect(
        service.registerUserMapping(mockTenantId, mockExternalUserId, 'unknown@contoso.com')
      ).rejects.toThrow('User not found in tenant');
    });

    it('should throw error when tenant not found', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { value: [mockGraphUser] },
      });
      tenantRepository.findOne.mockResolvedValueOnce(null);

      await expect(
        service.registerUserMapping(mockTenantId, mockExternalUserId, mockEmail)
      ).rejects.toThrow('Tenant not found or inactive');
    });

    it('should update existing mapping when external ID already mapped', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { value: [mockGraphUser] },
      });
      tenantRepository.findOne.mockResolvedValueOnce(mockTenant as MicrosoftTenant);
      tenantUserRepository.findOne.mockResolvedValueOnce(mockTenantUser as MicrosoftUser);
      tenantUserRepository.save.mockResolvedValueOnce(mockTenantUser as MicrosoftUser);

      await service.registerUserMapping(mockTenantId, mockExternalUserId, mockEmail);

      expect(tenantUserRepository.save).toHaveBeenCalledWith(
        expect.objectContaining({
          externalUserId: mockExternalUserId,
          microsoftUserId: mockMicrosoftUserId,
        })
      );
    });
  });

  describe('getMicrosoftUserId', () => {
    it('should get Microsoft user ID from database mapping', async () => {
      tenantRepository.findOne.mockResolvedValueOnce(mockTenant as MicrosoftTenant);
      tenantUserRepository.findOne.mockResolvedValueOnce(mockTenantUser as MicrosoftUser);

      const result = await service.getMicrosoftUserId(mockTenantId, mockExternalUserId);

      expect(result).toBe(mockMicrosoftUserId);
    });

    it('should return null when no mapping exists', async () => {
      tenantRepository.findOne.mockResolvedValueOnce(mockTenant as MicrosoftTenant);
      tenantUserRepository.findOne.mockResolvedValueOnce(null);

      const result = await service.getMicrosoftUserId(mockTenantId, 'unknown-user');

      expect(result).toBeNull();
    });

    it('should return null when tenant not found', async () => {
      tenantRepository.findOne.mockResolvedValueOnce(null);

      const result = await service.getMicrosoftUserId(mockTenantId, mockExternalUserId);

      expect(result).toBeNull();
    });

    it('should only return active mappings', async () => {
      tenantRepository.findOne.mockResolvedValueOnce(mockTenant as MicrosoftTenant);
      tenantUserRepository.findOne.mockResolvedValueOnce(null);

      await service.getMicrosoftUserId(mockTenantId, mockExternalUserId);

      expect(tenantUserRepository.findOne).toHaveBeenCalledWith(
        expect.objectContaining({
          where: expect.objectContaining({
            isActive: true,
          }),
        })
      );
    });
  });

  describe('listUsers', () => {
    it('should list users from tenant', async () => {
      const mockUsers = [mockGraphUser, { ...mockGraphUser, id: 'user-2', displayName: 'Jane Doe' }];
      mockedAxios.get.mockResolvedValueOnce({
        data: { value: mockUsers },
      });

      const result = await service.listUsers(mockTenantId);

      expect(result.users).toHaveLength(2);
      expect(result.users[0].microsoftUserId).toBe(mockMicrosoftUserId);
    });

    it('should support top parameter for pagination', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { value: [mockGraphUser] },
      });

      await service.listUsers(mockTenantId, { top: 50 });

      expect(mockedAxios.get).toHaveBeenCalledWith(
        expect.any(String),
        expect.objectContaining({
          params: expect.objectContaining({
            '$top': 50,
          }),
        })
      );
    });

    it('should support OData filter expressions', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { value: [mockGraphUser] },
      });

      await service.listUsers(mockTenantId, { filter: "accountEnabled eq true" });

      expect(mockedAxios.get).toHaveBeenCalledWith(
        expect.any(String),
        expect.objectContaining({
          params: expect.objectContaining({
            '$filter': "accountEnabled eq true",
          }),
        })
      );
    });

    it('should return nextLink for pagination', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: {
          value: [mockGraphUser],
          '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc123',
        },
      });

      const result = await service.listUsers(mockTenantId);

      expect(result.nextLink).toBe('https://graph.microsoft.com/v1.0/users?$skiptoken=abc123');
    });

    it('should support skipToken for continued pagination', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { value: [mockGraphUser] },
      });

      await service.listUsers(mockTenantId, { skipToken: 'abc123' });

      expect(mockedAxios.get).toHaveBeenCalledWith(
        expect.any(String),
        expect.objectContaining({
          params: expect.objectContaining({
            '$skiptoken': 'abc123',
          }),
        })
      );
    });

    it('should throw error on Graph API failure', async () => {
      mockedAxios.get.mockRejectedValueOnce(graphError('Forbidden'));

      await expect(
        service.listUsers(mockTenantId)
      ).rejects.toThrow('Failed to list users');
    });

    it('should return empty array when no users found', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { value: [] },
      });

      const result = await service.listUsers(mockTenantId);

      expect(result.users).toEqual([]);
    });
  });

  describe('caching', () => {
    it('should cache user lookups in memory', async () => {
      mockedAxios.get.mockResolvedValue({
        data: { value: [mockGraphUser] },
      });

      // First call
      await service.lookupUserByEmail(mockTenantId, mockEmail);

      // Second call should use cache
      await service.lookupUserByEmail(mockTenantId, mockEmail);

      expect(mockedAxios.get).toHaveBeenCalledTimes(1);
    });

    it('should clear cache when clearCache is called', async () => {
      mockedAxios.get.mockResolvedValue({
        data: { value: [mockGraphUser] },
      });

      // First call
      await service.lookupUserByEmail(mockTenantId, mockEmail);

      // Clear cache
      service.clearCache();

      // Second call should hit API
      await service.lookupUserByEmail(mockTenantId, mockEmail);

      expect(mockedAxios.get).toHaveBeenCalledTimes(2);
    });

    it('should cache lookups by both email and UPN', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { value: [mockGraphUser] },
      });

      // Lookup by email
      await service.lookupUserByEmail(mockTenantId, mockEmail);

      // Lookup by UPN should hit cache (since UPN was cached from email lookup)
      const result = await service.lookupUserByUpn(mockTenantId, mockUpn);

      expect(mockedAxios.get).toHaveBeenCalledTimes(1);
      expect(result?.microsoftUserId).toBe(mockMicrosoftUserId);
    });
  });

  describe('token acquisition', () => {
    it('should acquire app-only token for Graph calls', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { value: [mockGraphUser] },
      });

      await service.lookupUserByEmail(mockTenantId, mockEmail);

      expect(appOnlyAuthService.getAccessToken).toHaveBeenCalledWith(mockTenantId);
    });

    it('should include Bearer token in Authorization header', async () => {
      mockedAxios.get.mockResolvedValueOnce({
        data: { value: [mockGraphUser] },
      });

      await service.lookupUserByEmail(mockTenantId, mockEmail);

      expect(mockedAxios.get).toHaveBeenCalledWith(
        expect.any(String),
        expect.objectContaining({
          headers: expect.objectContaining({
            Authorization: `Bearer ${mockAccessToken}`,
          }),
        })
      );
    });
  });

  describe('error handling', () => {
    it('should throw descriptive error on lookup failure', async () => {
      mockedAxios.get.mockRejectedValueOnce(graphError('Network timeout'));

      await expect(
        service.lookupUserByEmail(mockTenantId, mockEmail)
      ).rejects.toThrow('Failed to lookup user by email: Network timeout');
    });

    it('should throw descriptive error when getting user by ID fails', async () => {
      mockedAxios.get.mockRejectedValueOnce(graphError('Server error'));

      await expect(
        service.getUserById(mockTenantId, mockMicrosoftUserId)
      ).rejects.toThrow('Failed to get user by ID: Server error');
    });

    it('should handle Graph API 403 error', async () => {
      mockedAxios.get.mockRejectedValueOnce(graphError('Forbidden', 403));

      await expect(
        service.listUsers(mockTenantId)
      ).rejects.toThrow('Failed to list users');
    });
  });
});
