import { Test, TestingModule } from '@nestjs/testing';
import { AppOnlyAuthService, AdminConsentResult } from './app-only-auth.service';
import { MICROSOFT_CONFIG } from '../../constants';
import { MicrosoftOutlookConfig } from '../../interfaces/config/outlook-config.interface';
import { MicrosoftTenant } from '../../entities/microsoft-tenant.entity';
import { MicrosoftTenantStatus } from '../../enums/microsoft-tenant-status.enum';
import axios from 'axios';
import * as fs from 'fs';

jest.mock('axios');
jest.mock('fs');
const mockedAxios = axios as jest.Mocked<typeof axios>;
const mockedFs = fs as jest.Mocked<typeof fs>;

describe('AppOnlyAuthService', () => {
  let service: AppOnlyAuthService;

  const mockTenantId = '12345678-1234-1234-1234-123456789abc';
  const mockClientId = 'test-client-id';
  const mockClientSecret = 'test-client-secret';
  const mockAccessToken = 'mock-access-token';
  const mockThumbprint = 'abc123def456ghi789jkl012mno345pqr678stu9';

  const createMockConfig = (overrides?: Partial<MicrosoftOutlookConfig>): MicrosoftOutlookConfig => ({
    clientId: mockClientId,
    clientSecret: mockClientSecret,
    redirectPath: '/auth/callback',
    backendBaseUrl: 'https://api.example.com',
    ...overrides,
  });

  const createModule = async (config: MicrosoftOutlookConfig): Promise<TestingModule> => {
    return Test.createTestingModule({
      providers: [
        AppOnlyAuthService,
        {
          provide: MICROSOFT_CONFIG,
          useValue: config,
        },
      ],
    }).compile();
  };

  const createMockTenant = (privateKey: string): MicrosoftTenant => {
    const tenant = new MicrosoftTenant();
    tenant.id = 1;
    tenant.tenantId = mockTenantId;
    tenant.clientId = 'tenant-client-id';
    tenant.certificateThumbprint = mockThumbprint;
    tenant.certificateKeyPath = '/path/to/tenant/key.pem';
    tenant.status = MicrosoftTenantStatus.ACTIVE;
    tenant.isActive = true;
    // Mock fs.readFileSync for this path
    mockedFs.readFileSync.mockImplementation((path: fs.PathOrFileDescriptor) => {
      if (path === '/path/to/tenant/key.pem') return privateKey;
      throw new Error('File not found');
    });
    return tenant;
  };

  beforeEach(() => {
    jest.clearAllMocks();
  });

  describe('initialization', () => {
    it('should initialize without app-only config', async () => {
      const module = await createModule(createMockConfig());
      service = module.get<AppOnlyAuthService>(AppOnlyAuthService);
      service.onModuleInit();

      expect(service).toBeDefined();
      expect(service.isEnabled()).toBe(false);
    });

    it('should initialize with app-only config disabled', async () => {
      const module = await createModule(createMockConfig({
        appOnly: {
          enabled: false,
          tenantId: mockTenantId,
        },
      }));
      service = module.get<AppOnlyAuthService>(AppOnlyAuthService);
      service.onModuleInit();

      expect(service.isEnabled()).toBe(false);
    });

    it('should initialize with app-only config enabled', async () => {
      const module = await createModule(createMockConfig({
        appOnly: {
          enabled: true,
          tenantId: mockTenantId,
        },
      }));
      service = module.get<AppOnlyAuthService>(AppOnlyAuthService);
      service.onModuleInit();

      expect(service.isEnabled()).toBe(true);
      expect(service.getTenantId()).toBe(mockTenantId);
    });

    it('should initialize with certificate from direct string', async () => {
      const { privateKey } = await generateTestKeyPair();

      const module = await createModule(createMockConfig({
        appOnly: {
          enabled: true,
          tenantId: mockTenantId,
          certificate: {
            privateKey,
            thumbprint: mockThumbprint,
          },
        },
      }));
      service = module.get<AppOnlyAuthService>(AppOnlyAuthService);
      service.onModuleInit();

      expect(service.isEnabled()).toBe(true);
    });

    it('should load certificate from file path', async () => {
      const { privateKey } = await generateTestKeyPair();
      mockedFs.readFileSync.mockReturnValue(privateKey);

      const module = await createModule(createMockConfig({
        appOnly: {
          enabled: true,
          tenantId: mockTenantId,
          certificate: {
            privateKeyPath: '/path/to/key.pem',
            thumbprint: mockThumbprint,
          },
        },
      }));
      service = module.get<AppOnlyAuthService>(AppOnlyAuthService);
      service.onModuleInit();

      expect(mockedFs.readFileSync).toHaveBeenCalledWith('/path/to/key.pem', 'utf8');
      expect(service.isEnabled()).toBe(true);
    });

    it('should load certificate from base64', async () => {
      const { privateKey } = await generateTestKeyPair();
      const base64Key = Buffer.from(privateKey).toString('base64');

      const module = await createModule(createMockConfig({
        appOnly: {
          enabled: true,
          tenantId: mockTenantId,
          certificate: {
            privateKeyBase64: base64Key,
            thumbprint: mockThumbprint,
          },
        },
      }));
      service = module.get<AppOnlyAuthService>(AppOnlyAuthService);
      service.onModuleInit();

      expect(service.isEnabled()).toBe(true);
    });

    it('should throw error when file path is invalid', async () => {
      mockedFs.readFileSync.mockImplementation(() => {
        throw new Error('ENOENT: no such file or directory');
      });

      const module = await createModule(createMockConfig({
        appOnly: {
          enabled: true,
          tenantId: mockTenantId,
          certificate: {
            privateKeyPath: '/invalid/path.pem',
            thumbprint: mockThumbprint,
          },
        },
      }));
      service = module.get<AppOnlyAuthService>(AppOnlyAuthService);

      expect(() => { service.onModuleInit(); }).toThrow(
        'Failed to load private key from file'
      );
    });
  });

  describe('isEnabled', () => {
    it('should return false when appOnly is not configured', async () => {
      const module = await createModule(createMockConfig());
      service = module.get<AppOnlyAuthService>(AppOnlyAuthService);

      expect(service.isEnabled()).toBe(false);
    });

    it('should return false when enabled is false', async () => {
      const module = await createModule(createMockConfig({
        appOnly: {
          enabled: false,
          tenantId: mockTenantId,
        },
      }));
      service = module.get<AppOnlyAuthService>(AppOnlyAuthService);

      expect(service.isEnabled()).toBe(false);
    });

    it('should return true when enabled is true and tenantId is set', async () => {
      const module = await createModule(createMockConfig({
        appOnly: {
          enabled: true,
          tenantId: mockTenantId,
        },
      }));
      service = module.get<AppOnlyAuthService>(AppOnlyAuthService);

      expect(service.isEnabled()).toBe(true);
    });
  });

  describe('getAccessToken with config', () => {
    let testPrivateKey: string;

    beforeAll(async () => {
      const keyPair = await generateTestKeyPair();
      testPrivateKey = keyPair.privateKey;
    });

    beforeEach(async () => {
      const module = await createModule(createMockConfig({
        appOnly: {
          enabled: true,
          tenantId: mockTenantId,
          certificate: {
            privateKey: testPrivateKey,
            thumbprint: mockThumbprint,
          },
        },
      }));
      service = module.get<AppOnlyAuthService>(AppOnlyAuthService);
      service.onModuleInit();
    });

    it('should request new token when cache is empty', async () => {
      mockedAxios.post.mockResolvedValueOnce({
        data: {
          access_token: mockAccessToken,
          token_type: 'Bearer',
          expires_in: 3600,
        },
      });

      const token = await service.getAccessToken();

      expect(token).toBe(mockAccessToken);
      expect(mockedAxios.post).toHaveBeenCalledWith(
        `https://login.microsoftonline.com/${mockTenantId}/oauth2/v2.0/token`,
        expect.any(String),
        expect.objectContaining({
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
          },
        })
      );
    });

    it('should return cached token when valid', async () => {
      mockedAxios.post.mockResolvedValueOnce({
        data: {
          access_token: mockAccessToken,
          token_type: 'Bearer',
          expires_in: 3600,
        },
      });

      // First call - fetches from API
      await service.getAccessToken();

      // Second call - should use cache
      const token = await service.getAccessToken();

      expect(token).toBe(mockAccessToken);
      expect(mockedAxios.post).toHaveBeenCalledTimes(1);
    });

    it('should use client assertion when certificate is configured', async () => {
      mockedAxios.post.mockResolvedValueOnce({
        data: {
          access_token: mockAccessToken,
          token_type: 'Bearer',
          expires_in: 3600,
        },
      });

      await service.getAccessToken();

      const postCall = mockedAxios.post.mock.calls[0];
      const requestBody = postCall[1] as string;
      expect(requestBody).toContain('client_assertion_type=');
      expect(requestBody).toContain('client_assertion=');
      expect(requestBody).not.toContain('client_secret=');
    });

    it('should throw error when tenant ID is not available', async () => {
      const module = await createModule(createMockConfig());
      const serviceWithoutTenant = module.get<AppOnlyAuthService>(AppOnlyAuthService);

      await expect(serviceWithoutTenant.getAccessToken()).rejects.toThrow(
        'App-only authentication requires a tenant ID'
      );
    });

    it('should throw error when token request fails', async () => {
      const axiosError = {
        isAxiosError: true,
        response: {
          status: 400,
          data: {
            error: 'invalid_client',
            error_description: 'Client authentication failed',
          },
        },
        message: 'Request failed',
      };
      mockedAxios.post.mockRejectedValueOnce(axiosError);
      (axios.isAxiosError as unknown as jest.Mock) = jest.fn().mockReturnValue(true);

      await expect(service.getAccessToken()).rejects.toThrow(
        'Failed to obtain app-only access token'
      );
    });
  });

  describe('getAccessToken with MicrosoftTenant entity', () => {
    let testPrivateKey: string;

    beforeAll(async () => {
      const keyPair = await generateTestKeyPair();
      testPrivateKey = keyPair.privateKey;
    });

    beforeEach(async () => {
      const module = await createModule(createMockConfig());
      service = module.get<AppOnlyAuthService>(AppOnlyAuthService);
      service.onModuleInit();
    });

    it('should get token using MicrosoftTenant entity credentials', async () => {
      const tenant = createMockTenant(testPrivateKey);

      mockedAxios.post.mockResolvedValueOnce({
        data: {
          access_token: mockAccessToken,
          token_type: 'Bearer',
          expires_in: 3600,
        },
      });

      const token = await service.getAccessToken(tenant);

      expect(token).toBe(mockAccessToken);
      expect(mockedAxios.post).toHaveBeenCalledWith(
        `https://login.microsoftonline.com/${mockTenantId}/oauth2/v2.0/token`,
        expect.stringContaining('client_id=tenant-client-id'),
        expect.any(Object)
      );
    });

    it('should use tenant clientId in JWT assertion', async () => {
      const tenant = createMockTenant(testPrivateKey);

      mockedAxios.post.mockResolvedValueOnce({
        data: {
          access_token: mockAccessToken,
          token_type: 'Bearer',
          expires_in: 3600,
        },
      });

      await service.getAccessToken(tenant);

      const postCall = mockedAxios.post.mock.calls[0];
      const requestBody = postCall[1] as string;

      // Extract client_assertion from request body
      const match = requestBody.match(/client_assertion=([^&]+)/);
      expect(match).toBeTruthy();

      const assertion = decodeURIComponent((match as RegExpMatchArray)[1]);
      const parts = assertion.split('.');
      const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString());

      expect(payload.iss).toBe('tenant-client-id');
      expect(payload.sub).toBe('tenant-client-id');
    });

    it('should cache private key loaded from entity', async () => {
      const tenant = createMockTenant(testPrivateKey);

      mockedAxios.post.mockResolvedValue({
        data: {
          access_token: mockAccessToken,
          token_type: 'Bearer',
          expires_in: 3600,
        },
      });

      // Clear all caches
      service.clearCache();
      jest.clearAllMocks();

      // First call - should read the private key file
      await service.getAccessToken(tenant);
      expect(mockedFs.readFileSync).toHaveBeenCalledTimes(1);

      // Clear only token cache (not private key cache)
      // Use a different tenant ID to not clear private key cache for original tenant
      service.invalidateCache('different-tenant');

      // Force token refresh by clearing this tenant's token cache directly
      // We need to make another token request
      mockedAxios.post.mockResolvedValue({
        data: {
          access_token: 'new-token',
          token_type: 'Bearer',
          expires_in: 3600,
        },
      });

      // Create a new tenant with same tenantId but mock won't be called again for privateKey
      const tenant2 = createMockTenant(testPrivateKey);
      jest.clearAllMocks(); // Clear the readFileSync calls from createMockTenant

      // The privateKey should still be cached
      await service.getAccessToken(tenant2);

      // Private key should NOT be read again (it's cached)
      expect(mockedFs.readFileSync).not.toHaveBeenCalled();
    });

    it('should throw error when tenant has no certificate thumbprint', async () => {
      const tenant = new MicrosoftTenant();
      tenant.tenantId = mockTenantId;
      tenant.clientId = 'test-client';
      tenant.certificateKeyPath = '/path/to/key.pem';
      // Missing certificateThumbprint

      await expect(service.getAccessToken(tenant)).rejects.toThrow(
        'has no certificate thumbprint configured'
      );
    });

    it('should throw error when tenant has no private key path', async () => {
      const tenant = new MicrosoftTenant();
      tenant.tenantId = mockTenantId;
      tenant.clientId = 'test-client';
      tenant.certificateThumbprint = mockThumbprint;
      // Missing certificateKeyPath

      await expect(service.getAccessToken(tenant)).rejects.toThrow(
        'has no private key path configured'
      );
    });
  });

  describe('buildClientAssertion', () => {
    it('should build valid JWT client assertion', async () => {
      const { privateKey } = await generateTestKeyPair();

      const module = await createModule(createMockConfig({
        appOnly: {
          enabled: true,
          tenantId: mockTenantId,
          certificate: {
            privateKey,
            thumbprint: mockThumbprint,
          },
        },
      }));
      service = module.get<AppOnlyAuthService>(AppOnlyAuthService);
      service.onModuleInit();

      const assertion = service.buildClientAssertion(mockTenantId);

      // JWT should have 3 parts
      const parts = assertion.split('.');
      expect(parts).toHaveLength(3);

      // Decode and verify header
      const header = JSON.parse(Buffer.from(parts[0], 'base64url').toString());
      expect(header.alg).toBe('PS256');
      expect(header.typ).toBe('JWT');
      expect(header['x5t#S256']).toBe(mockThumbprint);

      // Decode and verify payload
      const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString());
      expect(payload.iss).toBe(mockClientId);
      expect(payload.sub).toBe(mockClientId);
      expect(payload.aud).toContain(mockTenantId);
      expect(payload.jti).toBeDefined();
      expect(payload.exp).toBeGreaterThan(payload.iat);
    });
  });

  describe('getAdminConsentUrl', () => {
    beforeEach(async () => {
      const module = await createModule(createMockConfig({
        appOnly: {
          enabled: true,
          tenantId: mockTenantId,
        },
      }));
      service = module.get<AppOnlyAuthService>(AppOnlyAuthService);
      service.onModuleInit();
    });

    it('should generate admin consent URL for common endpoint', () => {
      const url = service.getAdminConsentUrl();

      expect(url).toContain('https://login.microsoftonline.com/common/adminconsent');
      expect(url).toContain(`client_id=${mockClientId}`);
      expect(url).toContain('redirect_uri=');
    });

    it('should include state parameter when provided', () => {
      const state = 'my-state-value';
      const url = service.getAdminConsentUrl(state);

      expect(url).toContain(`state=${state}`);
    });

    it('should use specified tenant ID', () => {
      const specificTenant = 'specific-tenant-id';
      const url = service.getAdminConsentUrl(undefined, specificTenant);

      expect(url).toContain(`https://login.microsoftonline.com/${specificTenant}/adminconsent`);
    });

    it('should use custom client ID when provided', () => {
      const customClientId = 'custom-client-id';
      const url = service.getAdminConsentUrl(undefined, 'common', customClientId);

      expect(url).toContain(`client_id=${customClientId}`);
    });
  });

  describe('handleAdminConsentCallback', () => {
    beforeEach(async () => {
      const module = await createModule(createMockConfig({
        appOnly: {
          enabled: true,
          tenantId: mockTenantId,
        },
      }));
      service = module.get<AppOnlyAuthService>(AppOnlyAuthService);
      service.onModuleInit();
    });

    it('should return success result when consent is granted', () => {
      const result: AdminConsentResult = service.handleAdminConsentCallback({
        tenant: mockTenantId,
        admin_consent: 'True',
        state: 'my-state',
      });

      expect(result.success).toBe(true);
      expect(result.tenantId).toBe(mockTenantId);
      expect(result.state).toBe('my-state');
      expect(result.error).toBeUndefined();
    });

    it('should return failure result when error is present', () => {
      const result: AdminConsentResult = service.handleAdminConsentCallback({
        tenant: mockTenantId,
        error: 'access_denied',
        error_description: 'The user denied consent',
        state: 'my-state',
      });

      expect(result.success).toBe(false);
      expect(result.tenantId).toBe(mockTenantId);
      expect(result.error).toBe('access_denied');
      expect(result.errorDescription).toBe('The user denied consent');
      expect(result.state).toBe('my-state');
    });

    it('should handle unexpected callback response', () => {
      const result: AdminConsentResult = service.handleAdminConsentCallback({
        // Missing required fields
      });

      expect(result.success).toBe(false);
      expect(result.error).toBe('unexpected_response');
    });
  });

  describe('cache management', () => {
    let testPrivateKey: string;

    beforeAll(async () => {
      const keyPair = await generateTestKeyPair();
      testPrivateKey = keyPair.privateKey;
    });

    beforeEach(async () => {
      const module = await createModule(createMockConfig({
        appOnly: {
          enabled: true,
          tenantId: mockTenantId,
          certificate: {
            privateKey: testPrivateKey,
            thumbprint: mockThumbprint,
          },
        },
      }));
      service = module.get<AppOnlyAuthService>(AppOnlyAuthService);
      service.onModuleInit();
    });

    it('should invalidate cache for specific tenant', async () => {
      mockedAxios.post.mockResolvedValue({
        data: {
          access_token: mockAccessToken,
          token_type: 'Bearer',
          expires_in: 3600,
        },
      });

      // Populate cache
      await service.getAccessToken();

      // Invalidate cache
      service.invalidateCache();

      // Should request new token
      await service.getAccessToken();

      expect(mockedAxios.post).toHaveBeenCalledTimes(2);
    });

    it('should clear all cached tokens', async () => {
      mockedAxios.post.mockResolvedValue({
        data: {
          access_token: mockAccessToken,
          token_type: 'Bearer',
          expires_in: 3600,
        },
      });

      // Populate cache
      await service.getAccessToken();

      // Clear all cache
      service.clearCache();

      // Should request new token
      await service.getAccessToken();

      expect(mockedAxios.post).toHaveBeenCalledTimes(2);
    });
  });

  describe('getTenantId', () => {
    it('should return configured tenant ID', async () => {
      const module = await createModule(createMockConfig({
        appOnly: {
          enabled: true,
          tenantId: mockTenantId,
        },
      }));
      service = module.get<AppOnlyAuthService>(AppOnlyAuthService);

      expect(service.getTenantId()).toBe(mockTenantId);
    });

    it('should return undefined when not configured', async () => {
      const module = await createModule(createMockConfig());
      service = module.get<AppOnlyAuthService>(AppOnlyAuthService);

      expect(service.getTenantId()).toBeUndefined();
    });
  });

  describe('getClientId', () => {
    it('should return the client ID', async () => {
      const module = await createModule(createMockConfig());
      service = module.get<AppOnlyAuthService>(AppOnlyAuthService);

      expect(service.getClientId()).toBe(mockClientId);
    });
  });
});

/**
 * Helper function to generate a test RSA key pair.
 */
async function generateTestKeyPair(): Promise<{ privateKey: string; publicKey: string }> {
  const crypto = await import('crypto');
  return new Promise((resolve, reject) => {
    crypto.generateKeyPair(
      'rsa',
      {
        modulusLength: 2048,
        publicKeyEncoding: { type: 'spki', format: 'pem' },
        privateKeyEncoding: { type: 'pkcs8', format: 'pem' },
      },
      (err, publicKey, privateKey) => {
        if (err) reject(err);
        else resolve({ publicKey, privateKey });
      }
    );
  });
}
