import { Controller, Get, Query, Logger, Res, HttpStatus } from '@nestjs/common';
import { Response } from 'express';
import { ApiTags, ApiResponse, ApiQuery, ApiOperation, ApiProduces } from '@nestjs/swagger';
import { MicrosoftAuthService } from '../services/auth/microsoft-auth.service';
import { MailboxInactiveError } from '../errors/mailbox-inactive.error';
import { CsrfValidationError } from '../errors/csrf-validation.error';
import { InvalidStateError } from '../errors/invalid-state.error';
import { SubscriptionSetupError, SubscriptionFailureReason } from '../errors/subscription-setup.error';

@ApiTags('Microsoft Auth')
@Controller('auth/microsoft')
export class MicrosoftAuthController {
  private readonly logger = new Logger(MicrosoftAuthController.name);

  constructor(private readonly microsoftAuthService: MicrosoftAuthService) {}

  /**
   * OAuth callback endpoint for Microsoft authentication
   *
   * @summary Process Microsoft OAuth callback
   * @description This endpoint handles the OAuth callback from Microsoft after a user has
   * authenticated with their Microsoft account. It exchanges the authorization code for
   * access and refresh tokens, saves them for the user, and sets up necessary webhooks
   * for calendar synchronization.
   *
   * The user ID is extracted from the state parameter that was passed during the initial
   * authorization request. The state parameter is base64 encoded and contains user ID and CSRF token.
   *
   * @param {string} code - The authorization code from Microsoft
   * @param {string} state - Base64 encoded state containing user ID and CSRF token
   * @returns {HTML} HTML page confirming successful authentication
   *
   * @throws {BadRequestException} When code or state is missing or invalid
   * @throws {InternalServerErrorException} When authentication fails
   */
  @Get('callback')
  @ApiOperation({
    summary: 'Microsoft OAuth callback handler',
    description:
      'Processes the callback from Microsoft OAuth authentication flow. Exchanges the authorization code for access and refresh tokens, saves them for the user, and sets up necessary webhooks for calendar synchronization. The user ID is extracted from the state parameter.',
  })
  @ApiQuery({
    name: 'code',
    description: 'Authorization code from Microsoft',
    required: true,
    type: String,
    example: 'M.R3_BAY.c0def4c2-12bf-0b29-9a3a-f8a1c4f56789',
  })
  @ApiQuery({
    name: 'state',
    description: 'Base64 encoded state containing user ID and CSRF token',
    required: true,
    type: String,
    example: 'eyJ1c2VySWQiOiI3IiwiY3NyZiI6IjEyMzQ1In0',
  })
  @ApiResponse({
    status: 200,
    description: 'Authentication successful, tokens saved and webhooks created',
    content: {
      'text/html': {
        example:
          '<h1>Authorization successful!</h1><p>Your Microsoft account has been linked successfully.</p>',
      },
    },
  })
  @ApiResponse({
    status: 400,
    description: 'Invalid or missing code/state parameters',
  })
  @ApiResponse({
    status: 500,
    description: 'Server error during authentication process',
  })
  @ApiProduces('text/html')
  async handleOauthCallback(
    @Query('code') code: string,
    @Query('state') state: string,
    @Res() res: Response,
  ) {
    try {
      if (!code || !state) {
        this.logger.error('Missing required parameters for OAuth callback');
        return res.status(HttpStatus.BAD_REQUEST).send('Missing required parameters');
      }

      // Exchange the code for tokens - no need to pass scopes as they'll be retrieved from state
      await this.microsoftAuthService.exchangeCodeForToken(code, state);

      // Return success message HTML
      return res.status(HttpStatus.OK).send(`
        <h1>Authorization successful!</h1>
        <p>Your Microsoft account has been linked successfully.</p>
        <p>You can close this tab now.</p>
        <script>
          // Optionally notify the parent window
          if (window.opener) {
            window.opener.postMessage('microsoft-auth-success', '*');
          }
        </script>
      `);
    } catch (error) {
      this.logger.error('Error handling OAuth callback:', error);

      // Malformed/truncated/missing state is bad client input (often a bot or
      // scanner replaying a chopped callback URL). Answer 400, not 500, so this
      // traffic stops paging as server errors.
      if (error instanceof InvalidStateError) {
        return res.status(HttpStatus.BAD_REQUEST).send('Invalid or malformed state parameter');
      }

      if (error instanceof CsrfValidationError) {
        return res.status(HttpStatus.OK).send(`
          <h1>Authorization Link Expired</h1>
          <p>This authorization link is no longer valid. This can happen if:</p>
          <ul>
            <li>The link was opened after it expired</li>
            <li>The page was refreshed after authorization completed</li>
            <li>The browser back button was used after completing authorization</li>
          </ul>
          <p>Please go back to the application and start the calendar connection again.</p>
          <p>You can close this tab now.</p>
          <script>
            if (window.opener) {
              window.opener.postMessage({ type: 'microsoft-auth-failed', error: 'Authorization link expired or already used. Please try connecting your calendar again.' }, '*');
            }
          </script>
        `);
      }

      if (error instanceof MailboxInactiveError) {
        return res.status(HttpStatus.OK).send(`
          <h1>Calendar Connection Failed</h1>
          <p>Your Microsoft account was authenticated, but we couldn't access your mailbox.</p>
          <p>This usually means your mailbox is either:</p>
          <ul>
            <li>Hosted on an on-premise Exchange server (not Exchange Online)</li>
            <li>Inactive or has been soft-deleted</li>
            <li>Missing an Exchange Online license</li>
          </ul>
          <p>Please contact your IT administrator to ensure your mailbox is hosted in Exchange Online (Microsoft 365).</p>
          <p>You can close this tab now.</p>
          <script>
            if (window.opener) {
              window.opener.postMessage({ type: 'microsoft-auth-failed', error: 'Your mailbox is not supported. It may be inactive, soft-deleted, or hosted on-premise. Please contact your IT administrator.' }, '*');
            }
          </script>
        `);
      }

      if (error instanceof SubscriptionSetupError) {
        const content = this.getSubscriptionErrorContent(error.reason);
        return res.status(HttpStatus.OK).send(`
          <h1>${content.title}</h1>
          ${content.body}
          <p>You can close this tab now.</p>
          <script>
            if (window.opener) {
              window.opener.postMessage({ type: 'microsoft-auth-failed', error: ${JSON.stringify(content.userError)} }, '*');
            }
          </script>
        `);
      }

      return res
        .status(HttpStatus.INTERNAL_SERVER_ERROR)
        .send('An error occurred during authentication');
    }
  }

  private getSubscriptionErrorContent(reason: SubscriptionFailureReason): {
    title: string;
    body: string;
    userError: string;
  } {
    switch (reason) {
      case SubscriptionFailureReason.PERMISSION_DENIED:
        return {
          title: 'Calendar Connection Failed',
          body: `
            <p>Your Microsoft account was authenticated, but your organization's settings prevent calendar notifications.</p>
            <p>This usually means:</p>
            <ul>
              <li>Your mailbox doesn't support this feature</li>
              <li>Your IT admin has restricted calendar access</li>
            </ul>
            <p>Please contact your IT administrator.</p>`,
          userError: 'Calendar setup blocked by organization policy. Please contact your IT administrator.',
        };
      case SubscriptionFailureReason.AUTH_EXPIRED:
        return {
          title: 'Authentication Error',
          body: '<p>Your authentication session expired during calendar setup.</p><p>Please try connecting your calendar again.</p>',
          userError: 'Authentication expired during setup. Please try again.',
        };
      case SubscriptionFailureReason.RATE_LIMITED:
        return {
          title: 'Calendar Setup Temporarily Unavailable',
          body: `<p>Microsoft's servers are currently busy.</p><p>Please wait a few minutes and try connecting your calendar again.</p>`,
          userError: 'Calendar setup is temporarily unavailable due to high demand. Please try again in a few minutes.',
        };
      case SubscriptionFailureReason.SERVICE_UNAVAILABLE:
        return {
          title: 'Microsoft Service Unavailable',
          body: `<p>Microsoft's calendar service is temporarily unavailable. This is not related to your account.</p><p>Please try again later.</p>`,
          userError: "Microsoft's service is temporarily unavailable. Please try again later.",
        };
      default:
        return {
          title: 'Calendar Subscription Failed',
          body: `<p>Your Microsoft account was authenticated, but we couldn't set up calendar notifications.</p><p>This is usually temporary. Please try connecting your calendar again.</p><p>If the problem persists, contact your administrator.</p>`,
          userError: 'Calendar notification setup failed. Please try connecting again.',
        };
    }
  }
}
