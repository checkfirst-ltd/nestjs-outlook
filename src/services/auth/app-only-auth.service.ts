import { Injectable, Logger, Inject, OnModuleInit } from '@nestjs/common';
import axios from 'axios';
import * as crypto from 'crypto';
import * as fs from 'fs';
import { MICROSOFT_CONFIG } from '../../constants';
import {
  MicrosoftOutlookConfig,
  AppOnlyAuthConfig,
  CertificateAuthConfig
} from '../../interfaces/config/outlook-config.interface';
import { MicrosoftTenant } from '../../entities/microsoft-tenant.entity';
import { TtlCache } from '../../utils/ttl-cache.util';

/**
 * Cached token entry with expiration tracking.
 */
interface CachedToken {
  accessToken: string;
  expiresAt: number;
}

/**
 * Microsoft token endpoint response for client credentials flow.
 */
interface ClientCredentialsTokenResponse {
  access_token: string;
  token_type: string;
  expires_in: number;
}

/**
 * Result of admin consent callback processing.
 */
export interface AdminConsentResult {
  /** The tenant ID that granted consent */
  tenantId: string;
  /** Whether admin consent was successful */
  success: boolean;
  /** Error message if consent failed */
  error?: string;
  /** Error description if consent failed */
  errorDescription?: string;
  /** State parameter passed through the consent flow */
  state?: string;
}

/**
 * Resolved tenant credentials for token requests.
 */
interface TenantCredentials {
  tenantId: string;
  clientId: string;
  privateKey: string;
  thumbprint: string;
}

/**
 * Service for app-only (client credentials) authentication with Microsoft Graph API.
 *
 * This service handles tenant-wide authentication without user delegation,
 * supporting both:
 * 1. Config-based authentication (single tenant from module config)
 * 2. Entity-based authentication (multiple tenants from MicrosoftTenant entities)
 *
 * Certificate authentication (recommended for production):
 * - Creates a signed JWT client assertion using the private key
 * - Uses x5t#S256 header with certificate thumbprint
 * - Signs with PS256 algorithm (RSA-PSS with SHA-256)
 *
 * Client secret authentication (simpler, less secure):
 * - Falls back to client_secret when certificate is not configured
 *
 * Token caching:
 * - Tokens are cached until 5 minutes before expiry
 * - Cache TTL is configurable via `tokenCacheTtlMs`
 *
 * Admin consent:
 * - Provides URL generation for admin consent flow
 * - Handles admin consent callback processing
 */
@Injectable()
export class AppOnlyAuthService implements OnModuleInit {
  private readonly logger = new Logger(AppOnlyAuthService.name);
  private readonly appOnlyConfig: AppOnlyAuthConfig | undefined;
  private readonly clientId: string;
  private readonly clientSecret: string;

  // Resolved private key from config (loaded from file, base64, or direct string)
  private resolvedPrivateKey: string | undefined;

  // Cache for private keys loaded from MicrosoftTenant entities (keyed by tenantId)
  private readonly privateKeyCache = new TtlCache<string, string>(300000); // 5 min TTL

  // Token endpoint template - tenant ID is inserted at runtime
  private readonly TOKEN_ENDPOINT_TEMPLATE = 'https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token';

  // Admin consent URL template
  private readonly ADMIN_CONSENT_URL_TEMPLATE = 'https://login.microsoftonline.com/{tenantId}/adminconsent';

  // Default scopes for Microsoft Graph API
  private readonly DEFAULT_SCOPES = ['https://graph.microsoft.com/.default'];

  // Default token cache TTL: 55 minutes (tokens typically expire in 1 hour)
  private readonly DEFAULT_TOKEN_CACHE_TTL_MS = 55 * 60 * 1000;

  // Buffer time before token expiry to trigger refresh (5 minutes)
  private readonly TOKEN_EXPIRY_BUFFER_MS = 5 * 60 * 1000;

  // JWT lifetime for client assertion (10 minutes max per Microsoft docs)
  private readonly CLIENT_ASSERTION_LIFETIME_S = 600;

  // Token cache keyed by tenant ID
  private readonly tokenCache: TtlCache<string, CachedToken>;

  constructor(
    @Inject(MICROSOFT_CONFIG)
    private readonly microsoftConfig: MicrosoftOutlookConfig,
  ) {
    this.appOnlyConfig = this.microsoftConfig.appOnly;
    this.clientId = this.microsoftConfig.clientId;
    this.clientSecret = this.microsoftConfig.clientSecret;

    // Initialize token cache with configured or default TTL
    const cacheTtl = this.appOnlyConfig?.tokenCacheTtlMs ?? this.DEFAULT_TOKEN_CACHE_TTL_MS;
    this.tokenCache = new TtlCache<string, CachedToken>(cacheTtl);
  }

  /**
   * Initialize the service - load private key from file/base64 if configured.
   */
  onModuleInit(): void {
    if (this.appOnlyConfig?.enabled && this.appOnlyConfig.certificate) {
      this.resolvedPrivateKey = this.resolvePrivateKey(this.appOnlyConfig.certificate);
      const authMethod = this.resolvedPrivateKey ? 'certificate' : 'client_secret';
      this.logger.log(
        `AppOnlyAuthService initialized for tenant ${this.appOnlyConfig.tenantId} using ${authMethod} authentication`
      );
    } else if (this.appOnlyConfig?.enabled) {
      this.logger.log(
        `AppOnlyAuthService initialized for tenant ${this.appOnlyConfig.tenantId} using client_secret authentication`
      );
    } else {
      this.logger.log('AppOnlyAuthService initialized (multi-tenant mode via entities)');
    }
  }

  /**
   * Resolve the private key from various sources.
   * Priority: privateKey > privateKeyBase64 > privateKeyPath — i.e. an env-provided key (direct
   * PEM or base64) always takes precedence over a file path, so the environment variable is
   * authoritative even when a legacy `privateKeyPath` is also present.
   */
  private resolvePrivateKey(certificate: CertificateAuthConfig): string | undefined {
    // Direct PEM string (env)
    if (certificate.privateKey) {
      this.logger.debug('Using private key from direct string');
      return certificate.privateKey;
    }

    // Base64-encoded (env) — wins over a file path.
    if (certificate.privateKeyBase64) {
      try {
        const key = Buffer.from(certificate.privateKeyBase64, 'base64').toString('utf8');
        this.logger.debug('Loaded private key from base64');
        return key;
      } catch (error) {
        this.logger.error(
          `Failed to decode base64 private key: ${error instanceof Error ? error.message : 'Unknown error'}`
        );
        throw new Error('Failed to decode base64 private key');
      }
    }

    // File path (fallback when no env-provided key material is configured).
    if (certificate.privateKeyPath) {
      return this.loadPrivateKeyFromFile(certificate.privateKeyPath);
    }

    return undefined;
  }

  /**
   * Load a private key from a file path.
   */
  private loadPrivateKeyFromFile(path: string): string {
    try {
      const key = fs.readFileSync(path, 'utf8');
      this.logger.debug(`Loaded private key from file: ${path}`);
      return key;
    } catch (error) {
      this.logger.error(
        `Failed to read private key from file ${path}: ${error instanceof Error ? error.message : 'Unknown error'}`
      );
      throw new Error(`Failed to load private key from file: ${path}`);
    }
  }

  /**
   * Check if app-only authentication is enabled and properly configured.
   */
  isEnabled(): boolean {
    return this.appOnlyConfig?.enabled === true && !!this.appOnlyConfig.tenantId;
  }

  /**
   * Get a valid access token for app-only (tenant-wide) operations.
   *
   * Supports two modes:
   * 1. Pass a MicrosoftTenant entity - uses entity's certificate paths and clientId
   * 2. Pass a tenant ID string - uses module config credentials
   *
   * Returns a cached token if still valid, otherwise requests a new one.
   * Tokens are refreshed 5 minutes before expiry to avoid request failures.
   *
   * @param tenantOrId - MicrosoftTenant entity or tenant ID string
   * @returns Promise resolving to a valid access token string
   * @throws Error if credentials are not available or token request fails
   */
  async getAccessToken(tenantOrId?: MicrosoftTenant | string): Promise<string> {
    // Resolve tenant credentials
    const credentials = this.resolveTenantCredentials(tenantOrId);

    // Check cache first
    const cached = this.tokenCache.get(credentials.tenantId);
    if (cached && !this.isTokenExpiringSoon(cached.expiresAt)) {
      this.logger.debug(`Using cached app-only token for tenant ${credentials.tenantId}`);
      return cached.accessToken;
    }

    // Request new token
    this.logger.debug(`Requesting new app-only token for tenant ${credentials.tenantId}`);
    return this.requestNewTokenWithCredentials(credentials);
  }

  /**
   * Resolve tenant credentials from various sources.
   */
  private resolveTenantCredentials(tenantOrId?: MicrosoftTenant | string): TenantCredentials {
    // MicrosoftTenant entity provided
    if (tenantOrId && typeof tenantOrId === 'object') {
      return this.resolveTenantEntityCredentials(tenantOrId);
    }

    // String tenant ID or undefined - use config
    const tenantId = tenantOrId ?? this.appOnlyConfig?.tenantId;

    if (!tenantId) {
      throw new Error('App-only authentication requires a tenant ID');
    }

    if (!this.isEnabled() && !tenantOrId) {
      throw new Error('App-only authentication is not enabled');
    }

    // Use config-based credentials
    if (this.resolvedPrivateKey && this.appOnlyConfig?.certificate?.thumbprint) {
      return {
        tenantId,
        clientId: this.clientId,
        privateKey: this.resolvedPrivateKey,
        thumbprint: this.appOnlyConfig.certificate.thumbprint,
      };
    }

    // No certificate - will fall back to client secret
    throw new Error('Certificate credentials required for app-only authentication');
  }

  /**
   * Resolve credentials from a MicrosoftTenant entity.
   *
   * The module-level key (the shared-app certificate, which may come from `certificate.privateKey`
   * / `privateKeyBase64` env config) ALWAYS takes precedence when configured — the environment
   * variable is authoritative and overrides any per-tenant `certificateKeyPath`. Only when no
   * module key is configured does it fall back to a per-tenant dedicated key file. The key and
   * thumbprint are always taken as a matching pair.
   */
  private resolveTenantEntityCredentials(tenant: MicrosoftTenant): TenantCredentials {
    // 1) Module-level (env-var / shared-app) key + thumbprint — authoritative when configured.
    if (this.resolvedPrivateKey && this.appOnlyConfig?.certificate?.thumbprint) {
      return {
        tenantId: tenant.tenantId,
        clientId: tenant.clientId,
        privateKey: this.resolvedPrivateKey,
        thumbprint: this.appOnlyConfig.certificate.thumbprint,
      };
    }

    // 2) Fall back to a per-tenant dedicated key file when no module key is configured.
    if (tenant.certificateKeyPath) {
      let privateKey = this.privateKeyCache.get(tenant.tenantId);
      if (!privateKey) {
        privateKey = this.loadPrivateKeyFromFile(tenant.certificateKeyPath);
        this.privateKeyCache.set(tenant.tenantId, privateKey);
      }
      const thumbprint = tenant.certificateThumbprint || this.appOnlyConfig?.certificate?.thumbprint;
      if (thumbprint) {
        return { tenantId: tenant.tenantId, clientId: tenant.clientId, privateKey, thumbprint };
      }
    }

    throw new Error(
      `Tenant ${tenant.tenantId} has no usable certificate: no module-level certificate key ` +
        `and no readable per-tenant key file.`
    );
  }

  /**
   * Request a new access token using resolved credentials.
   */
  private async requestNewTokenWithCredentials(credentials: TenantCredentials): Promise<string> {
    const tokenEndpoint = this.TOKEN_ENDPOINT_TEMPLATE.replace('{tenantId}', credentials.tenantId);
    const scopes = this.appOnlyConfig?.scopes ?? this.DEFAULT_SCOPES;

    // Build client assertion JWT
    const clientAssertion = this.buildClientAssertionWithCredentials(credentials);

    const params = new URLSearchParams({
      client_id: credentials.clientId,
      scope: scopes.join(' '),
      grant_type: 'client_credentials',
      client_assertion_type: 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer',
      client_assertion: clientAssertion,
    });

    try {
      const response = await axios.post<ClientCredentialsTokenResponse>(
        tokenEndpoint,
        params.toString(),
        {
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
          },
        }
      );

      const { access_token, expires_in } = response.data;
      const expiresAt = Date.now() + expires_in * 1000;

      // Cache the token
      this.tokenCache.set(credentials.tenantId, {
        accessToken: access_token,
        expiresAt,
      });

      this.logger.log(`Successfully obtained app-only token for tenant ${credentials.tenantId}, expires in ${String(expires_in)}s`);

      return access_token;
    } catch (error) {
      if (axios.isAxiosError(error)) {
        const errorData = error.response?.data as Record<string, unknown> | undefined;
        const errorCode = errorData?.error as string | undefined;
        const errorDescription = errorData?.error_description as string | undefined;

        // Surface the full Azure AD diagnostic payload. `error_codes` carries
        // the AADSTS numbers, and trace_id/correlation_id let you cross-reference
        // the failure in the Entra sign-in logs. These are the fields Microsoft
        // returns on a failed client-credentials token request.
        this.logger.error(
          `Failed to obtain app-only token for tenant ${credentials.tenantId}: ${errorCode ?? 'unknown'} - ${errorDescription ?? error.message}`,
          {
            status: error.response?.status,
            errorCode,
            errorDescription,
            errorCodes: errorData?.error_codes,
            traceId: errorData?.trace_id,
            correlationId: errorData?.correlation_id,
            timestamp: errorData?.timestamp,
            errorUri: errorData?.error_uri,
            tenantId: credentials.tenantId,
            tokenEndpoint,
          }
        );

        throw new Error(
          `Failed to obtain app-only access token: ${errorDescription ?? error.message}`
        );
      }

      this.logger.error(`Unexpected error requesting app-only token: ${error instanceof Error ? error.message : 'Unknown error'}`);
      throw error;
    }
  }

  /**
   * Build a signed JWT client assertion using explicit credentials.
   */
  private buildClientAssertionWithCredentials(credentials: TenantCredentials): string {
    const now = Math.floor(Date.now() / 1000);
    const tokenEndpoint = this.TOKEN_ENDPOINT_TEMPLATE.replace('{tenantId}', credentials.tenantId);

    // JWT Header with x5t#S256 (certificate thumbprint)
    const header = {
      alg: 'PS256',
      typ: 'JWT',
      'x5t#S256': credentials.thumbprint,
    };

    // JWT Payload
    const payload = {
      iss: credentials.clientId,
      sub: credentials.clientId,
      aud: tokenEndpoint,
      jti: crypto.randomUUID(),
      nbf: now,
      iat: now,
      exp: now + this.CLIENT_ASSERTION_LIFETIME_S,
    };

    // Sign the JWT
    return this.signJwt(header, payload, credentials.privateKey);
  }

  /**
   * Build a signed JWT client assertion for certificate-based authentication.
   * Uses config-based credentials.
   *
   * The assertion is a JWT signed with the private key, containing:
   * - Header: alg=PS256, typ=JWT, x5t#S256=<certificate thumbprint>
   * - Payload: iss=client_id, sub=client_id, aud=token_endpoint, jti=random, exp, iat, nbf
   *
   * @param tenantId - The Azure AD tenant ID
   * @returns Signed JWT string
   */
  buildClientAssertion(tenantId: string): string {
    if (!this.resolvedPrivateKey) {
      throw new Error('Private key is required for client assertion');
    }

    if (!this.appOnlyConfig?.certificate?.thumbprint) {
      throw new Error('Certificate thumbprint is required for client assertion');
    }

    return this.buildClientAssertionWithCredentials({
      tenantId,
      clientId: this.clientId,
      privateKey: this.resolvedPrivateKey,
      thumbprint: this.appOnlyConfig.certificate.thumbprint,
    });
  }

  /**
   * Generate the admin consent URL for a tenant administrator to grant
   * application permissions.
   *
   * The admin must visit this URL and sign in with their admin credentials
   * to grant consent for the application to access the tenant's resources.
   *
   * @param state - Optional state parameter to pass through the consent flow
   * @param tenantId - Optional tenant ID. Use 'common' for multi-tenant or 'organizations'
   * @param clientId - Optional client ID override (for per-tenant apps)
   * @returns The admin consent URL
   */
  getAdminConsentUrl(state?: string, tenantId: string = 'common', clientId?: string): string {
    const redirectUri = this.buildAdminConsentRedirectUri();
    const effectiveClientId = clientId ?? this.clientId;

    const params = new URLSearchParams({
      client_id: effectiveClientId,
      redirect_uri: redirectUri,
    });

    if (state) {
      params.append('state', state);
    }

    const consentUrl = this.ADMIN_CONSENT_URL_TEMPLATE.replace('{tenantId}', tenantId);
    return `${consentUrl}?${params.toString()}`;
  }

  /**
   * Handle the admin consent callback from Azure AD.
   *
   * After an admin grants consent, Azure AD redirects back to your application
   * with query parameters indicating success or failure.
   *
   * @param queryParams - The query parameters from the callback URL
   * @returns AdminConsentResult indicating success or failure
   */
  handleAdminConsentCallback(queryParams: {
    tenant?: string;
    admin_consent?: string;
    error?: string;
    error_description?: string;
    state?: string;
  }): AdminConsentResult {
    const { tenant, admin_consent, error, error_description, state } = queryParams;

    // Check for error response
    if (error) {
      this.logger.warn(
        `Admin consent failed: ${error} - ${error_description ?? 'No description'}`
      );
      return {
        tenantId: tenant ?? 'unknown',
        success: false,
        error,
        errorDescription: error_description,
        state,
      };
    }

    // Check for success response
    if (admin_consent === 'True' && tenant) {
      this.logger.log(`Admin consent granted for tenant ${tenant}`);
      return {
        tenantId: tenant,
        success: true,
        state,
      };
    }

    // Unexpected response
    this.logger.warn('Unexpected admin consent callback response', queryParams);
    return {
      tenantId: tenant ?? 'unknown',
      success: false,
      error: 'unexpected_response',
      errorDescription: 'Admin consent callback received unexpected parameters',
      state,
    };
  }

  /**
   * Build the redirect URI for admin consent callback.
   */
  private buildAdminConsentRedirectUri(): string {
    const config = this.microsoftConfig;

    if (!config.redirectPath) {
      throw new Error(
        'MicrosoftOutlookModule config is missing required field: redirectPath. ' +
        'Ensure it is provided when calling MicrosoftOutlookModule.forRoot() or MicrosoftOutlookModule.forRootAsync().',
      );
    }

    if (!config.backendBaseUrl) {
      throw new Error(
        'MicrosoftOutlookModule config is missing required field: backendBaseUrl. ' +
        'Ensure it is provided when calling MicrosoftOutlookModule.forRoot() or MicrosoftOutlookModule.forRootAsync().',
      );
    }

    // If redirectPath already contains a full URL, use it directly
    if (config.redirectPath.startsWith('http')) {
      // Replace the user auth callback path with admin consent path
      return config.redirectPath.replace(/\/callback$/, '/admin-consent/callback');
    }

    const baseUrl = config.backendBaseUrl.endsWith('/')
      ? config.backendBaseUrl.slice(0, -1)
      : config.backendBaseUrl;

    let path = '';

    if (config.basePath) {
      const cleanBasePath = config.basePath.replace(/^\/+|\/+$/g, '');
      path += `/${cleanBasePath}`;
    }

    // Use a dedicated admin consent callback path
    path += '/auth/microsoft/tenant/admin-callback';

    return `${baseUrl}${path}`;
  }

  /**
   * Sign a JWT using PS256 (RSA-PSS with SHA-256).
   *
   * @param header - JWT header object
   * @param payload - JWT payload object
   * @param privateKey - PEM-encoded private key
   * @returns Signed JWT string
   */
  private signJwt(
    header: Record<string, unknown>,
    payload: Record<string, unknown>,
    privateKey: string
  ): string {
    const encodedHeader = this.base64UrlEncode(JSON.stringify(header));
    const encodedPayload = this.base64UrlEncode(JSON.stringify(payload));
    const signingInput = `${encodedHeader}.${encodedPayload}`;

    // Create signature using PS256 (RSA-PSS with SHA-256)
    const sign = crypto.createSign('RSA-SHA256');
    sign.update(signingInput);

    const signature = sign.sign(
      {
        key: privateKey,
        padding: crypto.constants.RSA_PKCS1_PSS_PADDING,
        saltLength: crypto.constants.RSA_PSS_SALTLEN_DIGEST,
      },
      'base64'
    );

    const encodedSignature = this.base64ToBase64Url(signature);

    return `${signingInput}.${encodedSignature}`;
  }

  /**
   * Encode a string to Base64URL format.
   */
  private base64UrlEncode(str: string): string {
    return Buffer.from(str, 'utf8')
      .toString('base64')
      .replace(/\+/g, '-')
      .replace(/\//g, '_')
      .replace(/=/g, '');
  }

  /**
   * Convert standard Base64 to Base64URL format.
   */
  private base64ToBase64Url(base64: string): string {
    return base64
      .replace(/\+/g, '-')
      .replace(/\//g, '_')
      .replace(/=/g, '');
  }

  /**
   * Check if a token is expiring soon (within buffer time).
   */
  private isTokenExpiringSoon(expiresAt: number): boolean {
    return expiresAt - Date.now() < this.TOKEN_EXPIRY_BUFFER_MS;
  }

  /**
   * Invalidate cached token for a specific tenant.
   * Useful when token is known to be invalid and needs refresh.
   *
   * @param tenantId - The tenant ID to invalidate cache for
   */
  invalidateCache(tenantId?: string): void {
    const effectiveTenantId = tenantId ?? this.appOnlyConfig?.tenantId;
    if (effectiveTenantId) {
      this.tokenCache.delete(effectiveTenantId);
      this.privateKeyCache.delete(effectiveTenantId);
      this.logger.debug(`Invalidated token cache for tenant ${effectiveTenantId}`);
    }
  }

  /**
   * Clear all cached tokens.
   */
  clearCache(): void {
    this.tokenCache.clear();
    this.privateKeyCache.clear();
    this.logger.debug('Cleared all cached app-only tokens');
  }

  /**
   * Get the configured tenant ID.
   */
  getTenantId(): string | undefined {
    return this.appOnlyConfig?.tenantId;
  }

  /**
   * Get the client ID used for authentication.
   */
  getClientId(): string {
    return this.clientId;
  }

  /**
   * Get the module-level certificate configuration (thumbprint + file paths).
   *
   * Used when a tenant connection is recorded on the admin-consent callback and no
   * per-tenant certificate was pre-registered: the new row inherits the shared app
   * certificate so token acquisition (which resolves credentials from the entity)
   * can succeed. Returns `undefined` when the module is configured for client-secret
   * auth rather than a certificate.
   */
  getModuleCertificate():
    | { thumbprint: string; certificatePath?: string; privateKeyPath?: string }
    | undefined {
    const certificate = this.appOnlyConfig?.certificate;
    if (!certificate?.thumbprint) {
      return undefined;
    }
    return {
      thumbprint: certificate.thumbprint,
      certificatePath: certificate.certificatePath,
      privateKeyPath: certificate.privateKeyPath,
    };
  }
}
