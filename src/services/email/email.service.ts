import { Injectable, Logger, Inject } from '@nestjs/common';
import { Client } from '@microsoft/microsoft-graph-client';
import { MicrosoftAuthService } from '../auth/microsoft-auth.service';
import { MICROSOFT_CONFIG } from '../../constants';
import { MicrosoftOutlookConfig } from '../../interfaces/config/outlook-config.interface';
import { TokenResponse } from '../../interfaces/outlook/token-response.interface';
import { Message } from '@microsoft/microsoft-graph-types';

@Injectable()
export class EmailService {
  private readonly logger = new Logger(EmailService.name);

  constructor(
    private readonly microsoftAuthService: MicrosoftAuthService,
    @Inject(MICROSOFT_CONFIG)
    private readonly microsoftConfig: MicrosoftOutlookConfig,
  ) {}

  /**
   * Sends an email using Microsoft Graph API
   * 
   * @param message - The email message to send
   * @param accessToken - Access token for Microsoft Graph API
   * @param refreshToken - Refresh token for Microsoft Graph API
   * @param tokenExpiry - Expiry date of the access token
   * @param userId - User ID associated with the email account
   * @returns The sent message data and refreshed token data if tokens were refreshed
   */
  async sendEmail(
    message: Partial<Message>,
    accessToken: string,
    refreshToken: string,
    tokenExpiry: string | undefined,
    userId: number,
  ): Promise<{ message: Message; tokensRefreshed: boolean; refreshedTokens?: TokenResponse }> {
    try {
      let currentAccessToken = accessToken;
      let currentRefreshToken = refreshToken;
      let tokensRefreshed = false;
      let refreshedTokens: TokenResponse | undefined;

      // Check if token is expired and needs refresh
      if (currentRefreshToken && tokenExpiry) {
        if (this.microsoftAuthService.isTokenExpired(new Date(tokenExpiry))) {
          this.logger.log('Access token is expired or will expire soon. Refreshing token...');

          try {
            refreshedTokens = await this.microsoftAuthService.refreshAccessToken(
              currentRefreshToken,
              userId
            );

            // Update the access token
            currentAccessToken = refreshedTokens?.access_token || currentAccessToken;
            currentRefreshToken = refreshedTokens?.refresh_token || currentRefreshToken;
            tokensRefreshed = true;

            this.logger.log('Token refreshed successfully');
          } catch (refreshError) {
            this.logger.error('Failed to refresh token:', refreshError);
            throw new Error('Failed to refresh token');
          }
        }
      }

      // Initialize Microsoft Graph client with possibly refreshed token
      const client = Client.init({
        authProvider: (done) => {
          done(null, currentAccessToken);
        },
      });

      // Send the email
      const sentMessage = await client
        .api('/me/sendMail')
        .post({ message }) as Message;

      // Return both the message and token refresh information
      return {
        message: sentMessage,
        tokensRefreshed,
        refreshedTokens,
      };
    } catch (error: unknown) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to send email: ${errorMessage}`);
      throw new Error(`Failed to send email: ${errorMessage}`);
    }
  }
} 