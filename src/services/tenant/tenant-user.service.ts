import { Injectable, Logger, Inject } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { Repository, In } from 'typeorm';
import axios from 'axios';
import { MicrosoftTenant } from '../../entities/microsoft-tenant.entity';
import { MicrosoftUser } from '../../entities/microsoft-user.entity';
import { AppOnlyAuthService } from '../auth/app-only-auth.service';
import { TtlCache } from '../../utils/ttl-cache.util';
import { executeGraphApiCall } from '../../utils/outlook-api-executor.util';
import { mapWithConcurrency } from '../../utils/concurrent-map.util';
import { MICROSOFT_CONFIG } from '../../constants';
import { MicrosoftOutlookConfig } from '../../interfaces/config/outlook-config.interface';

/**
 * Microsoft Graph User object (simplified)
 */
interface GraphUser {
  id: string;
  userPrincipalName: string;
  displayName: string;
  mail: string | null;
}

/**
 * Result of a user lookup operation
 */
export interface TenantUserLookupResult {
  /** Microsoft Graph user ID (immutable ID) */
  microsoftUserId: string;
  /** User principal name (email-like identifier) */
  userPrincipalName: string;
  /** Display name */
  displayName: string;
  /** Primary email address */
  email: string | null;
}

/**
 * Summary of a tenant user-mapping teardown (see {@link TenantUserService.clearTenantUserMappings}).
 */
export interface ClearTenantMappingsResult {
  /** Rows that also had delegated OAuth tokens and were unmapped (app-only columns nulled, row kept). */
  delegatedRowsUnmapped: number;
  /** Pure app-only rows that were deleted. */
  appOnlyRowsDeleted: number;
  /** Delegated refresh tokens successfully revoked at Microsoft (only when `revokeDelegatedTokens` is set). */
  tokensRevoked: number;
  /** Delegated refresh-token revocations that failed (best-effort; teardown continues regardless). */
  tokenRevocationFailures: number;
}

/**
 * Service for looking up and mapping Microsoft 365 users within a tenant.
 *
 * This service provides tenant-wide user lookup capabilities using app-only
 * authentication. It maps external user identifiers (email, UPN) to Microsoft
 * Graph user IDs for use with other tenant services.
 *
 * Key features:
 * - Lookup users by email or user principal name
 * - Persist user mappings on the shared MicrosoftUser entity (one row per external user)
 * - Cache user ID mappings for performance
 * - Support for immutable IDs (IdType="ImmutableId")
 *
 * Required Graph API permissions (Application):
 * - User.Read.All
 */
@Injectable()
export class TenantUserService {
  private readonly logger = new Logger(TenantUserService.name);

  /**
   * Cache for user lookups: key = `${tenantId}:${identifier}`, value = microsoftUserId
   * TTL: 1 hour (user IDs don't change, but we want to handle deletions)
   */
  private readonly userIdCache = new TtlCache<string, TenantUserLookupResult>(60 * 60 * 1000);

  /**
   * Concurrency ceiling for per-user refresh-token revocation during teardown.
   * Kept small so a tenant with many mapped users doesn't burst Microsoft's
   * token endpoint (each revocation is an independent HTTP call).
   */
  private readonly REVOCATION_CONCURRENCY = 5;

  constructor(
    @InjectRepository(MicrosoftTenant)
    private readonly tenantRepository: Repository<MicrosoftTenant>,
    @InjectRepository(MicrosoftUser)
    private readonly tenantUserRepository: Repository<MicrosoftUser>,
    private readonly appOnlyAuthService: AppOnlyAuthService,
    @Inject(MICROSOFT_CONFIG)
    private readonly microsoftConfig: MicrosoftOutlookConfig,
  ) {}

  /**
   * Look up a Microsoft user by email address within a tenant.
   *
   * Uses the Microsoft Graph /users endpoint with $filter to find users
   * by their primary email or user principal name.
   *
   * @param tenantId - Microsoft tenant ID (Azure AD tenant GUID)
   * @param email - Email address to look up
   * @returns User lookup result with Microsoft user ID, or null if not found
   *
   * @example
   * ```typescript
   * const user = await tenantUserService.lookupUserByEmail(
   *   'tenant-guid-here',
   *   'john.doe@contoso.com'
   * );
   * if (user) {
   *   console.log(`Microsoft User ID: ${user.microsoftUserId}`);
   * }
   * ```
   */
  async lookupUserByEmail(
    tenantId: string,
    email: string,
  ): Promise<TenantUserLookupResult | null> {
    const cacheKey = `${tenantId}:email:${email.toLowerCase()}`;

    // Check cache first
    const cached = this.userIdCache.get(cacheKey);
    if (cached) {
      this.logger.debug(`[lookupUserByEmail] Cache hit for ${email} in tenant ${tenantId}`);
      return cached;
    }

    this.logger.log(`[lookupUserByEmail] Looking up user ${email} in tenant ${tenantId}`);

    try {
      // Get app-only access token for the tenant
      const accessToken = await this.appOnlyAuthService.getAccessToken(tenantId);

      // Use $filter to find user by mail or userPrincipalName
      // Note: mail property may be null for some users, UPN is always present.
      // Do NOT percent-encode the email here: it goes inside an OData string literal and axios
      // encodes the whole $filter param once. Pre-encoding turned '@' into '%40' inside the
      // literal, so `userPrincipalName eq 'user%40domain'` never matched the real UPN. Only
      // OData-escape single quotes (double them) per the OData literal spec.
      const escapedEmail = email.replace(/'/g, "''");
      const filterQuery = `mail eq '${escapedEmail}' or userPrincipalName eq '${escapedEmail}'`;

      const response = await executeGraphApiCall(
        () => axios.get<{ value: GraphUser[] }>(
          `https://graph.microsoft.com/v1.0/users`,
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
              'Content-Type': 'application/json',
              'Prefer': 'IdType="ImmutableId"',
            },
            params: {
              '$filter': filterQuery,
              '$select': 'id,userPrincipalName,displayName,mail',
              '$top': 1,
            },
          }
        ),
        {
          logger: this.logger,
          resourceName: `user lookup by email ${email}`,
          maxRetries: 3,
          return404AsNull: true,
        }
      );

      if (!response?.data.value || response.data.value.length === 0) {
        this.logger.warn(`[lookupUserByEmail] User not found: ${email} in tenant ${tenantId}`);
        return null;
      }

      const graphUser = response.data.value[0];
      const result: TenantUserLookupResult = {
        microsoftUserId: graphUser.id,
        userPrincipalName: graphUser.userPrincipalName,
        displayName: graphUser.displayName,
        email: graphUser.mail,
      };

      // Cache the result
      this.userIdCache.set(cacheKey, result);

      // Also cache by UPN for faster lookups
      const upnCacheKey = `${tenantId}:upn:${graphUser.userPrincipalName.toLowerCase()}`;
      this.userIdCache.set(upnCacheKey, result);

      this.logger.log(
        `[lookupUserByEmail] Found user ${email}: microsoftUserId=${graphUser.id}, UPN=${graphUser.userPrincipalName}`
      );

      return result;
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`[lookupUserByEmail] Failed to lookup user ${email}: ${errorMessage}`);
      throw new Error(`Failed to lookup user by email: ${errorMessage}`);
    }
  }

  /**
   * Look up a Microsoft user by their user principal name (UPN).
   *
   * UPN is typically in the format `user@domain.com` and is always present
   * for Microsoft 365 users.
   *
   * @param tenantId - Microsoft tenant ID (Azure AD tenant GUID)
   * @param upn - User principal name (e.g., john.doe@contoso.com)
   * @returns User lookup result with Microsoft user ID, or null if not found
   */
  async lookupUserByUpn(
    tenantId: string,
    upn: string,
  ): Promise<TenantUserLookupResult | null> {
    const cacheKey = `${tenantId}:upn:${upn.toLowerCase()}`;

    // Check cache first
    const cached = this.userIdCache.get(cacheKey);
    if (cached) {
      this.logger.debug(`[lookupUserByUpn] Cache hit for ${upn} in tenant ${tenantId}`);
      return cached;
    }

    this.logger.log(`[lookupUserByUpn] Looking up user by UPN ${upn} in tenant ${tenantId}`);

    try {
      // Get app-only access token for the tenant
      const accessToken = await this.appOnlyAuthService.getAccessToken(tenantId);

      // Direct lookup by UPN (faster than filter)
      const encodedUpn = encodeURIComponent(upn);

      const response = await executeGraphApiCall(
        () => axios.get<GraphUser>(
          `https://graph.microsoft.com/v1.0/users/${encodedUpn}`,
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
              'Content-Type': 'application/json',
              'Prefer': 'IdType="ImmutableId"',
            },
            params: {
              '$select': 'id,userPrincipalName,displayName,mail',
            },
          }
        ),
        {
          logger: this.logger,
          resourceName: `user lookup by UPN ${upn}`,
          maxRetries: 3,
          return404AsNull: true,
        }
      );

      if (!response?.data) {
        this.logger.warn(`[lookupUserByUpn] User not found: ${upn} in tenant ${tenantId}`);
        return null;
      }

      const graphUser = response.data;
      const result: TenantUserLookupResult = {
        microsoftUserId: graphUser.id,
        userPrincipalName: graphUser.userPrincipalName,
        displayName: graphUser.displayName,
        email: graphUser.mail,
      };

      // Cache the result
      this.userIdCache.set(cacheKey, result);

      // Also cache by email if present
      if (graphUser.mail) {
        const emailCacheKey = `${tenantId}:email:${graphUser.mail.toLowerCase()}`;
        this.userIdCache.set(emailCacheKey, result);
      }

      this.logger.log(
        `[lookupUserByUpn] Found user ${upn}: microsoftUserId=${graphUser.id}`
      );

      return result;
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`[lookupUserByUpn] Failed to lookup user ${upn}: ${errorMessage}`);
      throw new Error(`Failed to lookup user by UPN: ${errorMessage}`);
    }
  }

  /**
   * Get a Microsoft user by their Graph API user ID.
   *
   * Validates that a user ID exists and returns user details.
   * Useful for validating stored user IDs or getting fresh user info.
   *
   * @param tenantId - Microsoft tenant ID (Azure AD tenant GUID)
   * @param microsoftUserId - Microsoft Graph user ID
   * @returns User lookup result, or null if not found
   */
  async getUserById(
    tenantId: string,
    microsoftUserId: string,
  ): Promise<TenantUserLookupResult | null> {
    this.logger.log(`[getUserById] Getting user ${microsoftUserId} in tenant ${tenantId}`);

    try {
      // Get app-only access token for the tenant
      const accessToken = await this.appOnlyAuthService.getAccessToken(tenantId);

      const response = await executeGraphApiCall(
        () => axios.get<GraphUser>(
          `https://graph.microsoft.com/v1.0/users/${microsoftUserId}`,
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
              'Content-Type': 'application/json',
              'Prefer': 'IdType="ImmutableId"',
            },
            params: {
              '$select': 'id,userPrincipalName,displayName,mail',
            },
          }
        ),
        {
          logger: this.logger,
          resourceName: `user by ID ${microsoftUserId}`,
          maxRetries: 3,
          return404AsNull: true,
        }
      );

      if (!response?.data) {
        this.logger.warn(`[getUserById] User not found: ${microsoftUserId} in tenant ${tenantId}`);
        return null;
      }

      const graphUser = response.data;
      const result: TenantUserLookupResult = {
        microsoftUserId: graphUser.id,
        userPrincipalName: graphUser.userPrincipalName,
        displayName: graphUser.displayName,
        email: graphUser.mail,
      };

      // Cache the result for future lookups
      const upnCacheKey = `${tenantId}:upn:${graphUser.userPrincipalName.toLowerCase()}`;
      this.userIdCache.set(upnCacheKey, result);

      if (graphUser.mail) {
        const emailCacheKey = `${tenantId}:email:${graphUser.mail.toLowerCase()}`;
        this.userIdCache.set(emailCacheKey, result);
      }

      return result;
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`[getUserById] Failed to get user ${microsoftUserId}: ${errorMessage}`);
      throw new Error(`Failed to get user by ID: ${errorMessage}`);
    }
  }

  /**
   * Register a user mapping from external ID to Microsoft user ID.
   *
   * Upserts the shared MicrosoftUser row for this host user (keyed by externalUserId),
   * attaching the tenant + Microsoft identity. Reusing the existing row — rather than
   * creating a parallel one — keeps a single row per external user, so a user who also
   * completed delegated OAuth keeps their tokens on the same record and delegated
   * lookups by externalUserId stay unambiguous.
   *
   * @param tenantId - Microsoft tenant ID (Azure AD tenant GUID)
   * @param externalUserId - User ID from the host application
   * @param email - Email or UPN to look up the Microsoft user
   * @returns The created/updated user mapping
   */
  async registerUserMapping(
    tenantId: string,
    externalUserId: string,
    email: string,
  ): Promise<MicrosoftUser> {
    this.logger.log(
      `[registerUserMapping] Registering user mapping for ${externalUserId} -> ${email} in tenant ${tenantId}`
    );

    // Look up the Microsoft user
    const userLookup = await this.lookupUserByEmail(tenantId, email);
    if (!userLookup) {
      throw new Error(`User not found in tenant: ${email}`);
    }

    // Find the tenant entity
    const tenant = await this.tenantRepository.findOne({
      where: { tenantId, isActive: true },
    });

    if (!tenant) {
      throw new Error(`Tenant not found or inactive: ${tenantId}`);
    }

    // Upsert by externalUserId so we reuse an existing (possibly delegated) row rather
    // than creating a parallel record that would collide on externalUserId.
    let tenantUser = await this.tenantUserRepository.findOne({
      where: { externalUserId },
      relations: ['tenant'],
    });

    if (tenantUser) {
      // Attach/refresh the tenant identity on the existing row
      this.logger.log(`[registerUserMapping] Updating existing mapping for ${externalUserId}`);
      tenantUser.tenant = tenant;
      tenantUser.microsoftUserId = userLookup.microsoftUserId;
      tenantUser.userPrincipalName = userLookup.userPrincipalName;
      tenantUser.isActive = true;
    } else {
      // Create new mapping
      this.logger.log(`[registerUserMapping] Creating new mapping for ${externalUserId}`);
      tenantUser = new MicrosoftUser();
      tenantUser.tenant = tenant;
      tenantUser.externalUserId = externalUserId;
      tenantUser.microsoftUserId = userLookup.microsoftUserId;
      tenantUser.userPrincipalName = userLookup.userPrincipalName;
      tenantUser.isActive = true;
    }

    await this.tenantUserRepository.save(tenantUser);

    this.logger.log(
      `[registerUserMapping] Registered: ${externalUserId} -> ${userLookup.microsoftUserId} (${userLookup.userPrincipalName})`
    );

    return tenantUser;
  }

  /**
   * Bulk-load the persisted `MicrosoftUser` rows for a set of external user ids, in one (or a
   * few chunked) query. Used by bulk provisioning to detect which users already exist / are
   * already connected without issuing a query per user. The `IN (...)` list is chunked so a
   * very large batch can't exceed the database's parameter/packet limits.
   *
   * @param externalUserIds - Host user ids to look up.
   * @returns The existing rows; an external id with no row is simply absent from the result.
   */
  async findUsersByExternalIds(externalUserIds: string[]): Promise<MicrosoftUser[]> {
    if (externalUserIds.length === 0) {
      return [];
    }

    const CHUNK = 500;
    const found: MicrosoftUser[] = [];
    for (let i = 0; i < externalUserIds.length; i += CHUNK) {
      const chunk = externalUserIds.slice(i, i + CHUNK);
      const rows = await this.tenantUserRepository.find({
        where: { externalUserId: In(chunk) },
        // Load the tenant so callers can tell app-only users apart and read tenant status.
        relations: ['tenant'],
      });
      found.push(...rows);
    }
    return found;
  }

  /**
   * Get the Microsoft user ID for an external user ID.
   *
   * Looks up the mapping in the database first, then falls back to Graph API
   * lookup if not found.
   *
   * @param tenantId - Microsoft tenant ID (Azure AD tenant GUID)
   * @param externalUserId - User ID from the host application
   * @returns Microsoft user ID, or null if not found
   */
  async getMicrosoftUserId(
    tenantId: string,
    externalUserId: string,
  ): Promise<string | null> {
    // Check database mapping first
    const tenant = await this.tenantRepository.findOne({
      where: { tenantId, isActive: true },
    });

    if (!tenant) {
      this.logger.warn(`[getMicrosoftUserId] Tenant not found: ${tenantId}`);
      return null;
    }

    const tenantUser = await this.tenantUserRepository.findOne({
      where: {
        tenant: { id: tenant.id },
        externalUserId,
        isActive: true,
      },
    });

    if (tenantUser) {
      this.logger.debug(
        `[getMicrosoftUserId] Found mapping: ${externalUserId} -> ${tenantUser.microsoftUserId}`
      );
      return tenantUser.microsoftUserId;
    }

    this.logger.debug(
      `[getMicrosoftUserId] No mapping found for ${externalUserId} in tenant ${tenantId}`
    );
    return null;
  }

  /**
   * List users in a tenant with optional filtering.
   *
   * Useful for admin dashboards or bulk user operations.
   * Results are paginated by Microsoft Graph (default 100 per page).
   *
   * @param tenantId - Microsoft tenant ID (Azure AD tenant GUID)
   * @param options - Optional filtering and pagination options
   * @returns Array of user lookup results
   */
  async listUsers(
    tenantId: string,
    options?: {
      /** Maximum number of users to return (default: 100) */
      top?: number;
      /** Filter query (e.g., "accountEnabled eq true") */
      filter?: string;
      /** Skip token for pagination */
      skipToken?: string;
    },
  ): Promise<{ users: TenantUserLookupResult[]; nextLink?: string }> {
    this.logger.log(`[listUsers] Listing users in tenant ${tenantId}`);

    try {
      // Get app-only access token for the tenant
      const accessToken = await this.appOnlyAuthService.getAccessToken(tenantId);

      const params: Record<string, string | number> = {
        '$select': 'id,userPrincipalName,displayName,mail',
        '$top': options?.top ?? 100,
      };

      if (options?.filter) {
        params['$filter'] = options.filter;
      }

      if (options?.skipToken) {
        params['$skiptoken'] = options.skipToken;
      }

      const response = await executeGraphApiCall(
        () => axios.get<{ value: GraphUser[]; '@odata.nextLink'?: string }>(
          `https://graph.microsoft.com/v1.0/users`,
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
              'Content-Type': 'application/json',
              'Prefer': 'IdType="ImmutableId"',
            },
            params,
          }
        ),
        {
          logger: this.logger,
          resourceName: 'list users',
          maxRetries: 3,
        }
      );

      if (!response?.data.value) {
        return { users: [] };
      }

      const users: TenantUserLookupResult[] = response.data.value.map(graphUser => ({
        microsoftUserId: graphUser.id,
        userPrincipalName: graphUser.userPrincipalName,
        displayName: graphUser.displayName,
        email: graphUser.mail,
      }));

      this.logger.log(`[listUsers] Found ${users.length} users in tenant ${tenantId}`);

      return {
        users,
        nextLink: response.data['@odata.nextLink'],
      };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`[listUsers] Failed to list users: ${errorMessage}`);
      throw new Error(`Failed to list users: ${errorMessage}`);
    }
  }

  /**
   * Remove a tenant's app-only footprint from the shared `microsoft_users` table.
   *
   * Used by the tenant-disconnect flow. It runs at most two bulk SQL statements
   * (no per-row database loop, no recursion):
   *
   * - **Dual-capability rows** (a row that also carries delegated OAuth tokens): the
   *   app-only columns (`tenant_id`, `microsoft_user_id`, `user_principal_name`) are
   *   nulled in a single `UPDATE`, preserving the user's still-valid delegated login.
   * - **Pure app-only rows** (no delegated tokens): deleted in a single `DELETE`.
   *
   * When `revokeDelegatedTokens` is set, the teardown is aggressive: every delegated
   * refresh token for the tenant is first revoked at Microsoft (bounded concurrency to
   * respect rate limits), then **all** of the tenant's rows are deleted. Revocation is
   * best-effort — a failed revocation is counted but never aborts the teardown.
   *
   * @param tenantId - Microsoft tenant ID (Azure AD tenant GUID). Matched regardless of
   *   the tenant's `isActive` flag, so it works after the connection has been deactivated.
   * @param options.revokeDelegatedTokens - Also revoke and delete delegated rows (default: false).
   * @returns Counts of what was unmapped, deleted, and revoked.
   */
  async clearTenantUserMappings(
    tenantId: string,
    options?: { revokeDelegatedTokens?: boolean },
  ): Promise<ClearTenantMappingsResult> {
    const revokeDelegatedTokens = options?.revokeDelegatedTokens ?? false;
    const result: ClearTenantMappingsResult = {
      delegatedRowsUnmapped: 0,
      appOnlyRowsDeleted: 0,
      tokensRevoked: 0,
      tokenRevocationFailures: 0,
    };

    // Resolve the internal PK (the FK stored on microsoft_users.tenant_id). Match by GUID
    // only — disconnect deactivates the tenant first, so an isActive filter would miss it.
    const tenant = await this.tenantRepository.findOne({ where: { tenantId } });
    if (!tenant) {
      this.logger.warn(`[clearTenantUserMappings] Tenant not found: ${tenantId} — nothing to clear`);
      return result;
    }

    this.logger.log(
      `[clearTenantUserMappings] Clearing user mappings for tenant ${tenantId} ` +
      `(revokeDelegatedTokens=${revokeDelegatedTokens})`,
    );

    if (revokeDelegatedTokens) {
      // Aggressive purge: revoke every delegated refresh token, then delete all rows.
      // Bulk-select only the tokens we need (no entity hydration, no N+1).
      const rows = await this.tenantUserRepository
        .createQueryBuilder('u')
        .select('u.refresh_token', 'refreshToken')
        .where('u.tenant_id = :id', { id: tenant.id })
        .andWhere('u.refresh_token IS NOT NULL')
        .getRawMany<{ refreshToken: string | null }>();

      const refreshTokens = rows
        .map((row) => row.refreshToken)
        .filter((token): token is string => Boolean(token));

      if (refreshTokens.length > 0) {
        const revocations = await mapWithConcurrency(
          refreshTokens,
          this.REVOCATION_CONCURRENCY,
          (token) => this.revokeRefreshToken(token),
        );
        for (const ok of revocations) {
          if (ok) {
            result.tokensRevoked++;
          } else {
            result.tokenRevocationFailures++;
          }
        }
      }

      const deleteResult = await this.tenantUserRepository
        .createQueryBuilder()
        .delete()
        .where('tenant_id = :id', { id: tenant.id })
        .execute();
      result.appOnlyRowsDeleted = deleteResult.affected ?? 0;
    } else {
      // Safe default: keep delegated logins, drop only the app-only footprint.
      const now = new Date();

      const unmapResult = await this.tenantUserRepository
        .createQueryBuilder()
        .update()
        .set({
          tenant: null,
          microsoftUserId: null,
          userPrincipalName: null,
          updatedAt: now,
        })
        .where('tenant_id = :id', { id: tenant.id })
        .andWhere('refresh_token IS NOT NULL')
        .execute();
      result.delegatedRowsUnmapped = unmapResult.affected ?? 0;

      const deleteResult = await this.tenantUserRepository
        .createQueryBuilder()
        .delete()
        .where('tenant_id = :id', { id: tenant.id })
        .andWhere('refresh_token IS NULL')
        .execute();
      result.appOnlyRowsDeleted = deleteResult.affected ?? 0;
    }

    // Cached lookups for this tenant now point at gone/changed rows — drop them.
    this.userIdCache.clear();

    this.logger.log(
      `[clearTenantUserMappings] Done for tenant ${tenantId}: ` +
      `${result.delegatedRowsUnmapped} unmapped, ${result.appOnlyRowsDeleted} deleted, ` +
      `${result.tokensRevoked} tokens revoked, ${result.tokenRevocationFailures} revocation failures`,
    );

    return result;
  }

  /**
   * Revoke a single delegated refresh token at Microsoft. Best-effort: resolves `true`
   * on success and `false` on failure (logged, never thrown) so a bad token can't abort
   * a bulk teardown.
   */
  private async revokeRefreshToken(refreshToken: string): Promise<boolean> {
    try {
      await axios.post(
        'https://login.microsoftonline.com/common/oauth2/v2.0/logout',
        new URLSearchParams({
          token: refreshToken,
          token_type_hint: 'refresh_token',
        }),
        { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } },
      );
      return true;
    } catch (error) {
      this.logger.warn(
        `[revokeRefreshToken] Failed to revoke a delegated token: ${
          error instanceof Error ? error.message : 'Unknown error'
        }`,
      );
      return false;
    }
  }

  /**
   * Clear the user ID cache for a specific tenant.
   *
   * Useful when you know user data has changed or for testing.
   *
   * @param tenantId - Tenant ID to clear cache for (optional, clears all if not specified)
   */
  clearCache(tenantId?: string): void {
    if (tenantId) {
      this.logger.log(`[clearCache] Clearing user cache for tenant ${tenantId}`);
    } else {
      this.logger.log('[clearCache] Clearing entire user cache');
    }
    // Clear the entire cache - TtlCache will rebuild as needed
    this.userIdCache.clear();
  }
}
