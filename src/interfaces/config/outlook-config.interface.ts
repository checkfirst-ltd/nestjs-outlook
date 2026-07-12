import { RedisLike } from "../state/redis-like.interface";

/**
 * Optional shared-state backend configuration.
 *
 * When `redis.client` is provided, lock store and rate-limit store use Redis
 * so multiple containers coordinate. Otherwise both fall back to in-process
 * in-memory implementations (suitable only for single-container deployments).
 */
export interface MicrosoftOutlookStateConfig {
  redis?: {
    /** Host-provided ioredis-compatible client. nestjs-outlook never imports ioredis. */
    client: RedisLike;
    /** Key prefix applied to all keys. Defaults to "outlook:". */
    keyPrefix?: string;
    /**
     * If true, module init throws when the PING probe fails. Use in production
     * so a Redis outage triggers ECS restart instead of silently falling back
     * to in-memory (which re-introduces the cross-container concurrency bug).
     * Defaults to false.
     */
    required?: boolean;
  };
}

/**
 * Certificate-based authentication configuration.
 *
 * Uses a certificate (private key + thumbprint) to create a signed JWT
 * client assertion, which is more secure than client secrets.
 *
 * The certificate must be registered in Azure AD app registration
 * under "Certificates & secrets" > "Certificates".
 *
 * Supports multiple ways to provide the private key:
 * 1. Direct PEM string via `privateKey`
 * 2. File path via `privateKeyPath`
 * 3. Base64-encoded PEM via `privateKeyBase64`
 *
 * Priority: privateKey > privateKeyPath > privateKeyBase64
 */
export interface CertificateAuthConfig {
  /**
   * The certificate private key in PEM format (direct string).
   * This is used to sign the JWT client assertion.
   *
   * Example:
   * ```
   * -----BEGIN PRIVATE KEY-----
   * MIIEvgIBADANBg...
   * -----END PRIVATE KEY-----
   * ```
   */
  privateKey?: string;

  /**
   * Path to the private key file in PEM format.
   * Alternative to providing the key directly via `privateKey`.
   * The file is read at service initialization.
   */
  privateKeyPath?: string;

  /**
   * Base64-encoded private key in PEM format.
   * Alternative to providing the key directly via `privateKey`.
   * Useful for passing keys via environment variables.
   */
  privateKeyBase64?: string;

  /**
   * Path to the certificate file in PEM format.
   * Optional - only needed if you want to compute the thumbprint automatically.
   */
  certificatePath?: string;

  /**
   * Base64-encoded certificate in PEM format.
   * Optional - only needed if you want to compute the thumbprint automatically.
   */
  certificateBase64?: string;

  /**
   * The certificate thumbprint (SHA-256 hash of the certificate).
   * This is used in the x5t#S256 header of the JWT.
   *
   * Can be obtained from Azure portal under app registration
   * "Certificates & secrets" > "Certificates" > "Thumbprint",
   * or by computing SHA-256 hash of the certificate's DER encoding.
   *
   * Format: Base64url-encoded SHA-256 hash (43 characters)
   */
  thumbprint: string;
}

/**
 * Configuration for app-only (client credentials) authentication.
 *
 * This enables tenant-wide access without user delegation, using the
 * OAuth 2.0 client credentials flow. The app authenticates as itself
 * rather than on behalf of a user.
 *
 * Supports two authentication methods:
 * 1. Client secret (simpler, less secure)
 * 2. Certificate (more secure, recommended for production)
 *
 * If `certificate` is provided, it takes precedence over client secret.
 *
 * Required Microsoft Graph API permissions (Application type):
 * - Calendars.ReadWrite (read/write calendars for all users)
 * - Mail.ReadWrite (read/write mail for all users)
 * - User.Read.All (read user profiles)
 *
 * These permissions must be granted admin consent in Azure AD.
 */
export interface AppOnlyAuthConfig {
  /**
   * Enable app-only authentication mode.
   * When true, the module uses client credentials flow instead of
   * delegated user authentication.
   */
  enabled: boolean;

  /**
   * The Azure AD tenant ID for the organization.
   * This is required for app-only auth as the app authenticates
   * against a specific tenant rather than the common endpoint.
   * Format: GUID (e.g., "12345678-1234-1234-1234-123456789abc")
   */
  tenantId: string;

  /**
   * Optional: Certificate-based authentication configuration.
   * When provided, uses certificate (private key + thumbprint) to create
   * a signed JWT client assertion instead of using client secret.
   *
   * This is more secure than client secret and recommended for production.
   * The certificate must be registered in Azure AD.
   */
  certificate?: CertificateAuthConfig;

  /**
   * Optional: Override the default Microsoft Graph API scopes.
   * Defaults to ["https://graph.microsoft.com/.default"] which requests
   * all statically configured permissions for the app.
   */
  scopes?: string[];

  /**
   * Optional: Token cache TTL in milliseconds.
   * App-only tokens are cached to avoid unnecessary token requests.
   * Defaults to 3300000 (55 minutes) - slightly less than the typical
   * 1-hour token lifetime to ensure refresh before expiration.
   */
  tokenCacheTtlMs?: number;
}

/**
 * Configuration interface for Microsoft Outlook OAuth settings
 */
export interface MicrosoftOutlookConfig {
  /**
   * The client id for the Microsoft Outlook OAuth settings
   */
  clientId: string;
  /**
   * The client secret for the Microsoft Outlook OAuth settings
   */
  clientSecret: string;
  /**
   * The path of the redirect uri. e.g. auth/microsoft/callback
   */
  redirectPath: string;
  /**
   * The base url of the backend. e.g. https://dev.dashboard.checkfirstapp.com
   */
  backendBaseUrl: string;
  /**
   * The base path of the backend. e.g. api/v1
   */
  basePath?: string;
  /**
   * The path for the calendar webhook endpoint. e.g. /calendar/webhook
   * Defaults to /calendar/webhook
   */
  calendarWebhookPath?: string;
  /**
   * Optional shared-state backend. Use Redis to coordinate locks and
   * rate-limit budgets across multiple containers. Omit for in-memory.
   */
  state?: MicrosoftOutlookStateConfig;
  /**
   * TTL (ms) for the "revocation email already sent this cycle" dedupe flag,
   * which suppresses duplicate USER_REFRESH_TOKEN_INVALID emits when a burst of
   * webhooks arrives for a newly-revoked user.
   *
   * The flag is normally cleared when the user re-authenticates (account becomes
   * ACTIVE again). This TTL only bounds the case where the user never reconnects,
   * so the flag self-heals instead of living forever. Defaults to one week
   * (7 * 24 * 60 * 60 * 1000 ms). Lower it to re-notify sooner, or raise it to
   * suppress duplicate emails for longer.
   */
  revocationEmitFlagTtlMs?: number;
  /**
   * Optional app-only (client credentials) authentication configuration.
   * When configured and enabled, the module can authenticate as the application
   * itself for tenant-wide access without user delegation.
   *
   * This is useful for background services, scheduled tasks, or admin operations
   * that need to access resources across all users in a tenant.
   */
  appOnly?: AppOnlyAuthConfig;
}
