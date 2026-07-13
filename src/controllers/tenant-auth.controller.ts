import { Controller, Get, Delete, Query, Logger, Res, HttpStatus, Optional, Inject } from '@nestjs/common';
import { Response } from 'express';
import { ApiTags, ApiResponse, ApiQuery, ApiOperation, ApiProduces } from '@nestjs/swagger';
import { AppOnlyAuthService } from '../services/auth/app-only-auth.service';
import { MicrosoftTenantRepository } from '../repositories/microsoft-tenant.repository';
import { MicrosoftTenant } from '../entities/microsoft-tenant.entity';
import { MicrosoftTenantStatus } from '../enums/microsoft-tenant-status.enum';
import { TenantUserService, ClearTenantMappingsResult } from '../services/tenant/tenant-user.service';
import { MicrosoftSubscriptionService, BulkSubscriptionDeleteResult } from '../services/subscription/microsoft-subscription.service';

/**
 * Controller for handling tenant-wide (app-only) authentication flows.
 *
 * This controller manages the admin consent flow for enterprise tenants,
 * allowing organizations to grant tenant-wide access to their Microsoft 365
 * resources without requiring individual user authentication.
 *
 * The flow works as follows:
 * 1. Admin initiates consent by visiting the admin consent URL
 * 2. Microsoft redirects back to our callback endpoint
 * 3. We verify the consent was granted and activate the tenant connection
 * 4. The app can now access tenant resources using client credentials
 */
@ApiTags('Tenant Auth')
@Controller('auth/microsoft/tenant')
export class TenantAuthController {
  private readonly logger = new Logger(TenantAuthController.name);

  constructor(
    // NOTE: the `| null` union type makes TypeScript emit `Object` for this
    // parameter in `design:paramtypes`, so Nest cannot infer the provider token
    // by reflection. The explicit `@Inject(...)` token is required alongside
    // `@Optional()` — without it the dependency is silently injected as
    // `undefined`, making the controller always report "not configured".
    @Optional()
    @Inject(AppOnlyAuthService)
    private readonly appOnlyAuthService: AppOnlyAuthService | null,
    @Optional()
    @Inject(MicrosoftTenantRepository)
    private readonly tenantConnectionRepository: MicrosoftTenantRepository | null,
    // Optional teardown collaborators — only needed for the `purge` disconnect path.
    // Same `| null` reflection caveat as above: explicit @Inject token + @Optional.
    @Optional()
    @Inject(TenantUserService)
    private readonly tenantUserService: TenantUserService | null,
    @Optional()
    @Inject(MicrosoftSubscriptionService)
    private readonly subscriptionService: MicrosoftSubscriptionService | null,
  ) {}

  /**
   * Get the admin consent URL for initiating tenant-wide authentication.
   *
   * @summary Get admin consent URL
   * @description Returns the Microsoft admin consent URL that an Azure AD administrator
   * must visit to grant tenant-wide permissions to this application. The optional state
   * parameter can be used to correlate the consent flow with a specific tenant.
   *
   * @param {string} state - Optional state parameter to pass through the consent flow
   * @param {string} tenantId - Optional tenant ID to pre-select the tenant (default: 'common')
   * @returns {{ url: string }} Object containing the admin consent URL
   */
  @Get('admin-consent')
  @ApiOperation({
    summary: 'Get Microsoft admin consent URL',
    description:
      'Returns the admin consent URL for Azure AD administrators to grant tenant-wide permissions. The state parameter is passed through the consent flow for correlation.',
  })
  @ApiQuery({
    name: 'state',
    description: 'State parameter to pass through consent flow (e.g., external tenant ID)',
    required: false,
    type: String,
    example: 'my-org-tenant-001',
  })
  @ApiQuery({
    name: 'tenantId',
    description: 'Azure AD tenant ID to pre-select (default: common for any tenant)',
    required: false,
    type: String,
    example: '12345678-1234-1234-1234-123456789abc',
  })
  @ApiResponse({
    status: 200,
    description: 'Admin consent URL generated successfully',
    schema: {
      type: 'object',
      properties: {
        url: {
          type: 'string',
          example: 'https://login.microsoftonline.com/common/adminconsent?client_id=xxx&redirect_uri=xxx&state=xxx',
        },
      },
    },
  })
  @ApiResponse({
    status: 500,
    description: 'App-only authentication not configured',
  })
  getAdminConsentUrl(
    @Query('state') state?: string,
    @Query('tenantId') tenantId?: string,
  ): { url: string } {
    if (!this.appOnlyAuthService) {
      this.logger.error('App-only authentication is not configured');
      throw new Error('Tenant-wide authentication is not configured for this application');
    }

    const url = this.appOnlyAuthService.getAdminConsentUrl(
      state,
      tenantId ?? 'common',
    );

    this.logger.log(`Generated admin consent URL for state: ${state ?? 'none'}, tenant: ${tenantId ?? 'common'}`);

    return { url };
  }

  /**
   * Admin consent callback endpoint for tenant-wide authentication.
   *
   * @summary Process Microsoft admin consent callback
   * @description This endpoint handles the callback from Microsoft after a tenant
   * administrator has granted admin consent for the application. It verifies the
   * consent was successful and activates the tenant connection for app-only access.
   *
   * The state parameter contains the external tenant ID for correlation, and
   * the admin_consent parameter indicates whether consent was granted.
   *
   * @param {string} tenant - The Microsoft tenant ID from the consent flow
   * @param {string} state - External tenant ID passed through the consent flow
   * @param {string} admin_consent - Whether admin consent was granted ("True" or "False")
   * @param {string} error - Error code if consent was denied or failed
   * @param {string} error_description - Detailed error description
   * @returns {HTML} HTML page confirming consent status
   */
  @Get('admin-callback')
  @ApiOperation({
    summary: 'Microsoft admin consent callback handler',
    description:
      'Processes the callback from Microsoft admin consent flow. Verifies consent was granted and activates the tenant connection for app-only access to Microsoft 365 resources.',
  })
  @ApiQuery({
    name: 'tenant',
    description: 'Microsoft tenant ID (directory ID)',
    required: false,
    type: String,
    example: '12345678-1234-1234-1234-123456789abc',
  })
  @ApiQuery({
    name: 'state',
    description: 'External tenant ID for correlation',
    required: false,
    type: String,
    example: 'my-org-tenant-001',
  })
  @ApiQuery({
    name: 'admin_consent',
    description: 'Whether admin consent was granted',
    required: false,
    type: String,
    example: 'True',
  })
  @ApiQuery({
    name: 'error',
    description: 'Error code if consent failed',
    required: false,
    type: String,
  })
  @ApiQuery({
    name: 'error_description',
    description: 'Detailed error description',
    required: false,
    type: String,
  })
  @ApiResponse({
    status: 200,
    description: 'Consent flow completed (success or failure)',
    content: {
      'text/html': {
        example:
          '<h1>Admin Consent Granted!</h1><p>Your organization has been connected successfully.</p>',
      },
    },
  })
  @ApiProduces('text/html')
  async handleAdminConsentCallback(
    @Query('tenant') microsoftTenantId: string,
    @Query('state') externalTenantId: string,
    @Query('admin_consent') adminConsent: string,
    @Query('error') error: string,
    @Query('error_description') errorDescription: string,
    @Res() res: Response,
  ) {
    const correlationId = `tenant-consent-${externalTenantId || 'unknown'}-${Date.now()}`;

    // Observability: always record exactly what Microsoft returned on the
    // admin-consent callback, before any branching. These are the raw query
    // parameters Microsoft appends to the redirect URI.
    this.logger.log(
      `[${correlationId}] Admin consent callback received from Microsoft: ` +
        JSON.stringify({
          tenant: microsoftTenantId,
          state: externalTenantId,
          admin_consent: adminConsent,
          error: error,
          error_description: errorDescription,
        }),
    );

    try {
      // Check if app-only auth is configured
      if (!this.appOnlyAuthService || !this.tenantConnectionRepository) {
        this.logger.error(`[${correlationId}] App-only authentication is not configured`);
        return res.status(HttpStatus.OK).send(this.renderErrorPage(
          'Configuration Error',
          'Tenant-wide authentication is not configured for this application.',
          'Please contact your administrator to enable enterprise tenant connections.',
        ));
      }

      // Handle consent denied or error cases
      if (error) {
        this.logger.warn(
          `[${correlationId}] Admin consent denied or failed: ${error} - ${errorDescription}`
        );

        // Update tenant status if we have the external ID
        if (externalTenantId) {
          await this.tenantConnectionRepository.updateStatus(
            externalTenantId,
            MicrosoftTenantStatus.CONSENT_REVOKED
          );
        }

        return res.status(HttpStatus.OK).send(this.renderConsentDeniedPage(error, errorDescription));
      }

      // Validate required parameters
      if (!microsoftTenantId || !externalTenantId) {
        this.logger.error(`[${correlationId}] Missing required parameters`);
        return res.status(HttpStatus.OK).send(this.renderErrorPage(
          'Invalid Request',
          'The consent callback is missing required parameters.',
          'Please try the admin consent process again.',
        ));
      }

      // Verify consent was granted
      if (!adminConsent || adminConsent.toLowerCase() !== 'true') {
        this.logger.warn(`[${correlationId}] Admin consent not granted: ${adminConsent}`);
        return res.status(HttpStatus.OK).send(this.renderConsentDeniedPage(
          'consent_not_granted',
          'Administrator did not grant consent for the requested permissions.'
        ));
      }

      this.logger.log(
        `[${correlationId}] Admin consent granted for tenant ${microsoftTenantId} (external: ${externalTenantId})`
      );

      // Look up an existing connection for this tenant. It may not exist yet:
      // callers can initiate admin consent by simply supplying a tenant ID, without
      // pre-registering a MicrosoftTenant row. In that case we record the tenant here,
      // after consent has been confirmed, rather than failing.
      let connection = await this.tenantConnectionRepository.findByExternalTenantId(externalTenantId);

      if (!connection) {
        this.logger.log(
          `[${correlationId}] No existing tenant connection for ${microsoftTenantId}; recording a new one from admin-consent callback`
        );

        // Inherit the module-level (shared app) certificate so app-only token
        // acquisition can resolve credentials from the new row below. Configuration
        // Error is only reachable if the module has no certificate configured.
        const moduleCertificate = this.appOnlyAuthService.getModuleCertificate();
        if (!moduleCertificate) {
          this.logger.error(
            `[${correlationId}] Cannot record tenant ${microsoftTenantId}: no module certificate configured`
          );
          return res.status(HttpStatus.OK).send(this.renderErrorPage(
            'Configuration Error',
            'Admin consent was granted, but the application has no certificate configured to complete the connection.',
            'Please contact your administrator to configure the application certificate.',
          ));
        }

        connection = new MicrosoftTenant();
        connection.tenantId = microsoftTenantId;
        connection.clientId = this.appOnlyAuthService.getClientId();
        connection.certificateThumbprint = moduleCertificate.thumbprint;
        connection.certificatePath = moduleCertificate.certificatePath ?? null;
        connection.certificateKeyPath = moduleCertificate.privateKeyPath ?? null;
      }

      // Verify the Microsoft tenant ID matches
      if (connection.tenantId && connection.tenantId !== microsoftTenantId) {
        this.logger.error(
          `[${correlationId}] Microsoft tenant ID mismatch: expected ${connection.tenantId}, got ${microsoftTenantId}`
        );
        return res.status(HttpStatus.OK).send(this.renderErrorPage(
          'Tenant Mismatch',
          'The Microsoft tenant ID does not match the expected value.',
          'Please ensure you are granting consent from the correct Azure AD tenant.',
        ));
      }

      // Update tenant connection with Microsoft tenant ID and activate
      connection.tenantId = microsoftTenantId;
      connection.status = MicrosoftTenantStatus.ACTIVE;
      connection.isActive = true;
      connection.adminConsentGrantedAt = new Date();
      connection = await this.tenantConnectionRepository.save(connection);

      // Test that we can obtain a token for this tenant.
      // Pass the tenant ENTITY (not the tenant-ID string) so credentials are
      // resolved from the registered row (per-tenant certificate/key path).
      // Passing a bare string would route to the module-level config instead,
      // which fails with "Certificate credentials required" when only
      // per-tenant credentials are configured.
      try {
        await this.appOnlyAuthService.getAccessToken(connection);
        this.logger.log(`[${correlationId}] Successfully verified token acquisition for tenant ${microsoftTenantId}`);
      } catch (tokenError) {
        const tokenErrorMsg = tokenError instanceof Error ? tokenError.message : 'Unknown error';
        this.logger.error(`[${correlationId}] Failed to acquire token after consent: ${tokenErrorMsg}`);

        // Update status to indicate issue
        await this.tenantConnectionRepository.updateStatus(
          externalTenantId,
          MicrosoftTenantStatus.PENDING_CONSENT
        );

        return res.status(HttpStatus.OK).send(this.renderErrorPage(
          'Token Acquisition Failed',
          'Admin consent was granted, but we could not verify access.',
          'This may be a temporary issue. Please try again in a few minutes, or contact your administrator if the problem persists.',
        ));
      }

      this.logger.log(`[${correlationId}] Tenant connection activated successfully`);

      // Return success page
      return res.status(HttpStatus.OK).send(this.renderSuccessPage(externalTenantId));

    } catch (error) {
      this.logger.error(`[${correlationId}] Error handling admin consent callback:`, error);
      return res.status(HttpStatus.OK).send(this.renderErrorPage(
        'Unexpected Error',
        'An unexpected error occurred while processing the admin consent.',
        'Please try again. If the problem persists, contact support.',
      ));
    }
  }

  /**
   * Get the current tenant connection status.
   *
   * @summary Get tenant connection
   * @description Returns the stored tenant connection for the given tenant, or for the
   * module-configured tenant when no `tenantId` is supplied. Returns `null` when no active
   * connection exists (never connected, or disconnected), which callers read as "not connected".
   *
   * @param {string} tenantId - Optional Azure AD tenant ID; defaults to the configured tenant
   * @returns {MicrosoftTenant | null} The tenant connection record, or null
   */
  @Get('connection')
  @ApiOperation({
    summary: 'Get tenant connection status',
    description:
      'Returns the stored app-only tenant connection (status, consent timestamp, active flag). Falls back to the module-configured tenant when tenantId is omitted. Returns null when no active connection exists.',
  })
  @ApiQuery({
    name: 'tenantId',
    description: 'Azure AD tenant ID to look up (defaults to the configured tenant)',
    required: false,
    type: String,
    example: '12345678-1234-1234-1234-123456789abc',
  })
  @ApiResponse({
    status: 200,
    description: 'Tenant connection record, or null when not connected',
  })
  async getConnection(@Query('tenantId') tenantId?: string) {
    if (!this.appOnlyAuthService || !this.tenantConnectionRepository) {
      return null;
    }

    // An explicit tenantId is looked up directly (and only that one).
    // findByTenantId only returns active rows, so a disconnected tenant reads as null.
    if (tenantId) {
      return (await this.tenantConnectionRepository.findByTenantId(tenantId)) ?? null;
    }

    // No tenantId supplied: prefer the module-configured tenant when it is a concrete
    // tenant. When it is 'common' (or has no row), this is the dynamic-tenant flow where
    // the tenant is chosen at consent time — fall back to the single active connection so
    // callers (e.g. the dashboard) reflect it without knowing the tenant id up front.
    const configuredTenantId = this.appOnlyAuthService.getTenantId();
    if (configuredTenantId && configuredTenantId !== 'common') {
      const byConfigured = await this.tenantConnectionRepository.findByTenantId(configuredTenantId);
      if (byConfigured) {
        return byConfigured;
      }
    }

    const activeConnections = await this.tenantConnectionRepository.findAllActive();
    return activeConnections[0] ?? null;
  }

  /**
   * Disconnect a tenant connection.
   *
   * @summary Disconnect tenant
   * @description Deactivates the stored tenant connection and invalidates any cached
   * app-only access token so subsequent Graph calls stop working until re-consent.
   *
   * By default this is a **soft** disconnect: the tenant row is flagged inactive and its
   * token cache dropped, but the mapped `microsoft_users` rows and any Outlook webhook
   * subscriptions are left in place. Pass `purge=true` for a **full teardown** that also
   * deletes the tenant's Outlook subscriptions at Microsoft and clears its user mappings.
   * Add `revokeUserTokens=true` to additionally revoke and remove delegated user tokens.
   *
   * @param {string} tenantId - Optional Azure AD tenant ID; defaults to the configured tenant
   * @param {boolean} purge - Also delete Outlook subscriptions and clear user mappings
   * @param {boolean} revokeUserTokens - With purge, also revoke/remove delegated user tokens
   * @returns Confirmation message, plus a teardown summary when purging
   */
  @Delete('connection')
  @ApiOperation({
    summary: 'Disconnect the tenant connection',
    description:
      'Deactivates the stored app-only tenant connection and invalidates the cached access token. ' +
      'With purge=true, also deletes the tenant\'s Outlook webhook subscriptions at Microsoft (rate-limited) ' +
      'and clears its user mappings; with revokeUserTokens=true it additionally revokes delegated tokens. ' +
      'Falls back to the module-configured tenant when tenantId is omitted.',
  })
  @ApiQuery({
    name: 'tenantId',
    description: 'Azure AD tenant ID to disconnect (defaults to the configured tenant)',
    required: false,
    type: String,
    example: '12345678-1234-1234-1234-123456789abc',
  })
  @ApiQuery({
    name: 'purge',
    description: 'Also delete Outlook subscriptions and clear user mappings for the tenant',
    required: false,
    type: Boolean,
    example: true,
  })
  @ApiQuery({
    name: 'revokeUserTokens',
    description: 'With purge, revoke and remove delegated user tokens (implies purge)',
    required: false,
    type: Boolean,
    example: false,
  })
  @ApiResponse({
    status: 200,
    description: 'Tenant disconnected (or nothing to disconnect)',
  })
  async disconnect(
    @Query('tenantId') tenantId?: string,
    @Query('purge') purge?: string | boolean,
    @Query('revokeUserTokens') revokeUserTokens?: string | boolean,
  ): Promise<{
    message: string;
    subscriptions?: BulkSubscriptionDeleteResult;
    userMappings?: ClearTenantMappingsResult;
  }> {
    if (!this.appOnlyAuthService || !this.tenantConnectionRepository) {
      throw new Error('Tenant-wide authentication is not configured for this application');
    }

    // Query params arrive as strings ('true'/'false'); coerce leniently. revokeUserTokens
    // implies purge — you can't revoke tokens without tearing down the mappings.
    const revokeTokens = this.isTruthyFlag(revokeUserTokens);
    const shouldPurge = this.isTruthyFlag(purge) || revokeTokens;

    // Resolve which tenant to disconnect, mirroring getConnection: an explicit tenantId
    // wins; otherwise use the module-configured tenant when concrete; otherwise (the
    // 'common' dynamic-tenant flow) fall back to the active connection. Without this
    // fallback a no-tenantId disconnect would deactivate 'common' — a no-op — and the
    // real connection would stay active.
    let resolvedTenantId = tenantId;
    if (!resolvedTenantId) {
      const configuredTenantId = this.appOnlyAuthService.getTenantId();
      if (
        configuredTenantId &&
        configuredTenantId !== 'common' &&
        (await this.tenantConnectionRepository.findByTenantId(configuredTenantId))
      ) {
        resolvedTenantId = configuredTenantId;
      } else {
        const activeConnections = await this.tenantConnectionRepository.findAllActive();
        resolvedTenantId = activeConnections[0]?.tenantId;
      }
    }

    if (!resolvedTenantId) {
      return { message: 'No tenant connection to disconnect.' };
    }

    const response: {
      message: string;
      subscriptions?: BulkSubscriptionDeleteResult;
      userMappings?: ClearTenantMappingsResult;
    } = { message: 'Microsoft 365 tenant disconnected successfully.' };

    // Full teardown runs BEFORE deactivation so subscription deletion still has a valid
    // app-only token (findByTenantId only returns active rows once deactivated).
    if (shouldPurge) {
      if (this.subscriptionService) {
        response.subscriptions =
          await this.subscriptionService.deleteAllAppOnlySubscriptionsForTenant(resolvedTenantId);
      } else {
        this.logger.warn(
          `[disconnect] purge requested but subscription service is unavailable; ` +
          `skipping Outlook subscription cleanup for ${resolvedTenantId}`,
        );
      }

      if (this.tenantUserService) {
        response.userMappings = await this.tenantUserService.clearTenantUserMappings(
          resolvedTenantId,
          { revokeDelegatedTokens: revokeTokens },
        );
      } else {
        this.logger.warn(
          `[disconnect] purge requested but tenant user service is unavailable; ` +
          `skipping user-mapping cleanup for ${resolvedTenantId}`,
        );
      }
    }

    await this.tenantConnectionRepository.deactivate(resolvedTenantId);
    this.appOnlyAuthService.invalidateCache(resolvedTenantId);
    this.logger.log(
      `Tenant connection disconnected: ${resolvedTenantId} (purge=${shouldPurge}, revokeUserTokens=${revokeTokens})`,
    );

    return response;
  }

  /**
   * Coerce a query-string flag to boolean. Query params are strings, so treat
   * 'true'/'1'/'yes' (case-insensitive) and boolean `true` as truthy; everything
   * else (including undefined and 'false') is false.
   */
  private isTruthyFlag(value: string | boolean | undefined): boolean {
    if (typeof value === 'boolean') {
      return value;
    }
    if (typeof value === 'string') {
      const normalized = value.trim().toLowerCase();
      return normalized === 'true' || normalized === '1' || normalized === 'yes';
    }
    return false;
  }

  /**
   * Render the success HTML page after admin consent is granted.
   */
  private renderSuccessPage(externalTenantId: string): string {
    return `
      <!DOCTYPE html>
      <html lang="en">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Admin Consent Granted</title>
        <style>
          body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; max-width: 600px; margin: 50px auto; padding: 20px; text-align: center; }
          h1 { color: #107c10; }
          .success-icon { font-size: 64px; margin-bottom: 20px; }
          .tenant-id { background: #f0f0f0; padding: 10px; border-radius: 4px; font-family: monospace; margin: 20px 0; }
          .close-hint { color: #666; margin-top: 30px; }
        </style>
      </head>
      <body>
        <div class="success-icon">&#10003;</div>
        <h1>Admin Consent Granted!</h1>
        <p>Your organization has been successfully connected.</p>
        <div class="tenant-id">Tenant: ${this.escapeHtml(externalTenantId)}</div>
        <p>The application now has tenant-wide access to the approved Microsoft 365 resources.</p>
        <p class="close-hint">You can close this tab now.</p>
        <script>
          if (window.opener) {
            window.opener.postMessage({ type: 'tenant-consent-success', tenantId: '${this.escapeJs(externalTenantId)}' }, '*');
          }
        </script>
      </body>
      </html>
    `;
  }

  /**
   * Render the consent denied HTML page.
   */
  private renderConsentDeniedPage(error: string, description: string): string {
    const errorMessages: Record<string, { title: string; explanation: string }> = {
      'access_denied': {
        title: 'Access Denied',
        explanation: 'The administrator declined to grant the requested permissions.',
      },
      'consent_required': {
        title: 'Consent Required',
        explanation: 'Administrator consent is required for the requested permissions.',
      },
      'consent_not_granted': {
        title: 'Consent Not Granted',
        explanation: 'The administrator did not grant consent for the application.',
      },
    };

    const knownError = error in errorMessages ? errorMessages[error as 'access_denied' | 'consent_required' | 'consent_not_granted'] : null;
    const errorInfo = knownError ?? {
      title: 'Consent Failed',
      explanation: description || 'An error occurred during the consent process.',
    };

    return `
      <!DOCTYPE html>
      <html lang="en">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Admin Consent Not Granted</title>
        <style>
          body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; max-width: 600px; margin: 50px auto; padding: 20px; text-align: center; }
          h1 { color: #d83b01; }
          .error-icon { font-size: 64px; margin-bottom: 20px; }
          .error-details { background: #fef0f0; border: 1px solid #fdd; padding: 15px; border-radius: 4px; margin: 20px 0; text-align: left; }
          .close-hint { color: #666; margin-top: 30px; }
        </style>
      </head>
      <body>
        <div class="error-icon">&#10007;</div>
        <h1>${this.escapeHtml(errorInfo.title)}</h1>
        <p>${this.escapeHtml(errorInfo.explanation)}</p>
        ${description ? `<div class="error-details"><strong>Details:</strong> ${this.escapeHtml(description)}</div>` : ''}
        <p>To connect your organization, an Azure AD administrator must grant consent for the required permissions.</p>
        <p class="close-hint">You can close this tab now.</p>
        <script>
          if (window.opener) {
            window.opener.postMessage({ type: 'tenant-consent-failed', error: '${this.escapeJs(error)}' }, '*');
          }
        </script>
      </body>
      </html>
    `;
  }

  /**
   * Render a generic error HTML page.
   */
  private renderErrorPage(title: string, message: string, action: string): string {
    return `
      <!DOCTYPE html>
      <html lang="en">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>${this.escapeHtml(title)}</title>
        <style>
          body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; max-width: 600px; margin: 50px auto; padding: 20px; text-align: center; }
          h1 { color: #d83b01; }
          .error-icon { font-size: 64px; margin-bottom: 20px; }
          .action { background: #f5f5f5; padding: 15px; border-radius: 4px; margin: 20px 0; }
          .close-hint { color: #666; margin-top: 30px; }
        </style>
      </head>
      <body>
        <div class="error-icon">&#9888;</div>
        <h1>${this.escapeHtml(title)}</h1>
        <p>${this.escapeHtml(message)}</p>
        <div class="action">${this.escapeHtml(action)}</div>
        <p class="close-hint">You can close this tab now.</p>
        <script>
          if (window.opener) {
            window.opener.postMessage({ type: 'tenant-consent-error', error: '${this.escapeJs(title)}' }, '*');
          }
        </script>
      </body>
      </html>
    `;
  }

  /**
   * Escape HTML special characters to prevent XSS.
   */
  private escapeHtml(str: string): string {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  /**
   * Escape string for use in JavaScript.
   */
  private escapeJs(str: string): string {
    return str
      .replace(/\\/g, '\\\\')
      .replace(/'/g, "\\'")
      .replace(/"/g, '\\"')
      .replace(/\n/g, '\\n')
      .replace(/\r/g, '\\r');
  }
}
