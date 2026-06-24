import { Controller, Get, Query, Logger, Res, HttpStatus, Optional } from '@nestjs/common';
import { Response } from 'express';
import { ApiTags, ApiResponse, ApiQuery, ApiOperation, ApiProduces } from '@nestjs/swagger';
import { AppOnlyAuthService } from '../services/auth/app-only-auth.service';
import { MicrosoftTenantRepository } from '../repositories/microsoft-tenant.repository';
import { MicrosoftTenantStatus } from '../enums/microsoft-tenant-status.enum';

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
    @Optional()
    private readonly appOnlyAuthService: AppOnlyAuthService | null,
    @Optional()
    private readonly tenantConnectionRepository: MicrosoftTenantRepository | null,
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

      // Update tenant connection status
      const connection = await this.tenantConnectionRepository.findByExternalTenantId(externalTenantId);

      if (!connection) {
        this.logger.error(`[${correlationId}] Tenant connection not found: ${externalTenantId}`);
        return res.status(HttpStatus.OK).send(this.renderErrorPage(
          'Tenant Not Found',
          'The tenant connection was not found in our system.',
          'Please ensure the tenant was properly registered before requesting admin consent.',
        ));
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
      await this.tenantConnectionRepository.save(connection);

      // Test that we can obtain a token for this tenant
      try {
        await this.appOnlyAuthService.getAccessToken(microsoftTenantId);
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
