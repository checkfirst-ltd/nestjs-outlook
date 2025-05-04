import { Injectable, Logger, Inject, forwardRef } from '@nestjs/common';
import { EventEmitter2 } from '@nestjs/event-emitter';
import axios from 'axios';
import { TokenResponse } from '../interfaces/outlook/token-response.interface';
import * as qs from 'querystring';
import { OutlookService } from './outlook.service';
import { Subscription } from '@microsoft/microsoft-graph-types';
import { MICROSOFT_CONFIG } from '../constants';
import { MicrosoftOutlookConfig } from '../interfaces/config/outlook-config.interface';
import { OutlookEventTypes } from '../event-types.enum';
import * as crypto from 'crypto';
import { Cron, CronExpression } from '@nestjs/schedule';
import { MicrosoftCsrfTokenRepository } from '../repositories/microsoft-csrf-token.repository';
import { MicrosoftTokenApiResponse } from '../interfaces/microsoft-auth/microsoft-token-api-response.interface';
import { StateObject } from '../interfaces/microsoft-auth/state-object.interface';

@Injectable()
export class MicrosoftAuthService {
  private readonly logger = new Logger(MicrosoftAuthService.name);
  private readonly clientId: string;
  private readonly clientSecret: string;
  private readonly tenantId = 'common';
  private readonly redirectUri: string;
  private readonly tokenEndpoint: string;
  private readonly scope: string;
  // CSRF token expiration time (30 minutes)
  private readonly CSRF_TOKEN_EXPIRY = 30 * 60 * 1000;

  constructor(
    private readonly eventEmitter: EventEmitter2,
    @Inject(forwardRef(() => OutlookService))
    private readonly outlookService: OutlookService,
    @Inject(MICROSOFT_CONFIG)
    private readonly microsoftConfig: MicrosoftOutlookConfig,
    private readonly csrfTokenRepository: MicrosoftCsrfTokenRepository,
  ) {
    console.log('MicrosoftAuthService constructor - microsoftConfig:', {
      clientId: this.microsoftConfig.clientId,
      redirectUri: this.microsoftConfig.redirectPath,
    });

    this.clientId = this.microsoftConfig.clientId;
    this.clientSecret = this.microsoftConfig.clientSecret;

    // Build the redirect URI based on config
    this.redirectUri = this.buildRedirectUri(this.microsoftConfig);
    console.log('Redirect URI:', this.redirectUri);
    this.tokenEndpoint = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
    this.scope = ['offline_access', 'Calendars.ReadWrite', 'Calendars.Read', 'User.Read'].join(' ');

    // Log the redirect URI to help with debugging
    this.logger.log(`Microsoft OAuth redirect URI set to: ${this.redirectUri}`);
  }

  /**
   * Builds redirect URI based on configuration values
   * Format: backendBaseUrl/[basePath]/redirectPath
   * @param config Microsoft Outlook Configuration
   * @returns Complete redirect URI
   */
  private buildRedirectUri(config: MicrosoftOutlookConfig): string {
    // If redirectPath already contains a full URL, use it directly
    if (config.redirectPath && config.redirectPath.startsWith('http')) {
      this.logger.log(`Using complete redirect URI from config: ${config.redirectPath}`);
      return config.redirectPath;
    }

    // If no backendBaseUrl is provided, use default localhost
    const baseUrl = config.backendBaseUrl || 'http://localhost:3000';

    // Remove trailing slash from baseUrl if exists
    const cleanBaseUrl = baseUrl.endsWith('/') ? baseUrl.slice(0, -1) : baseUrl;

    // Build path components
    let path = '';

    // Add basePath if it exists
    if (config.basePath) {
      // Remove leading and trailing slashes, then add a single leading slash
      const cleanBasePath = config.basePath.replace(/^\/+|\/+$/g, '');
      path += `/${cleanBasePath}`;
    }

    // Add redirectPath (removing leading slash if it exists)
    if (config.redirectPath) {
      const cleanRedirectPath = config.redirectPath.replace(/^\/+/g, '');
      path += `/${cleanRedirectPath}`;
    } else {
      // Fallback to default redirect path if none provided
      path += '/api/v1/auth/microsoft/callback';
    }

    // Ensure the path doesn't have double slashes
    path = path.replace(/\/+/g, '/');

    const finalUri = `${cleanBaseUrl}${path}`;
    this.logger.log(`Constructed redirect URI: ${finalUri}`);
    this.logger.debug(
      `Using config: baseUrl=${baseUrl}, basePath=${config.basePath || ''}, redirectPath=${config.redirectPath || ''}`,
    );

    return finalUri;
  }

  /**
   * Scheduled job to clean up expired tokens
   * Runs every day at midnight
   */
  @Cron(CronExpression.EVERY_DAY_AT_MIDNIGHT)
  async cleanupExpiredTokens() {
    try {
      await this.csrfTokenRepository.cleanupExpiredTokens();
      this.logger.log('Cleaned up expired CSRF tokens');
    } catch (error) {
      this.logger.error(
        `Error cleaning up expired tokens: ${error instanceof Error ? error.message : 'Unknown error'}`,
      );
    }
  }

  /**
   * Generate a secure CSRF token
   */
  private async generateCsrfToken(userId: string | number): Promise<string> {
    const token = crypto.randomBytes(32).toString('hex');

    // Save token in the database
    await this.csrfTokenRepository.saveToken(token, userId, this.CSRF_TOKEN_EXPIRY);

    return token;
  }

  /**
   * Parse state parameter from OAuth callback
   */
  public parseState(state: string): StateObject | null {
    try {
      // Add padding back if needed
      const paddingNeeded = 4 - (state.length % 4);
      const paddedState = paddingNeeded < 4 ? state + '='.repeat(paddingNeeded) : state;

      const decoded = Buffer.from(paddedState, 'base64').toString();
      return JSON.parse(decoded) as StateObject;
    } catch (error) {
      this.logger.error(
        `Failed to parse state: ${error instanceof Error ? error.message : 'Unknown error'}`,
      );
      return null;
    }
  }

  /**
   * Verify a CSRF token and its associated timestamp
   * @param token The CSRF token to validate
   * @param timestamp Optional timestamp for additional expiration check
   * @returns String error message if validation fails, null if token is valid
   */
  public async validateCsrfToken(token: string, timestamp?: number): Promise<string | null> {
    // Check if token exists
    if (!token) {
      return 'Missing CSRF token';
    }

    // Find and validate token from database
    const csrfToken = await this.csrfTokenRepository.findAndValidateToken(token);

    // If token doesn't exist or has expired
    if (!csrfToken) {
      this.logger.warn('CSRF token not found or expired');
      return 'Invalid or expired CSRF token';
    }

    // If timestamp is provided, validate it as well
    if (timestamp && Date.now() - timestamp > this.CSRF_TOKEN_EXPIRY) {
      this.logger.warn(`Request timestamp expired for user ${csrfToken.userId}`);
      return 'Authorization request has expired';
    }

    // Token is valid
    return null;
  }

  /**
   * Get the Microsoft login URL
   */
  async getLoginUrl(userId: string): Promise<string> {
    // Generate a secure CSRF token linked to this user
    const csrf = await this.generateCsrfToken(userId);

    // Generate state with user ID and CSRF token
    const stateObj = {
      userId,
      csrf,
      timestamp: Date.now(),
    };
    const stateJson = JSON.stringify(stateObj);
    const state = Buffer.from(stateJson).toString('base64').replace(/=/g, ''); // Remove padding '=' characters

    this.logger.log(`State object: ${JSON.stringify(stateObj)}`);

    // Ensure proper URI encoding for parameters
    const encodedRedirectUri = encodeURIComponent(this.redirectUri);
    const encodedScope = encodeURIComponent(this.scope);

    this.logger.log(`Redirect URI (raw): ${this.redirectUri}`);
    this.logger.log(`Redirect URI (encoded): ${encodedRedirectUri}`);

    const authorizeUrl =
      `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/authorize` +
      `?client_id=${this.clientId}` +
      `&response_type=code` +
      `&redirect_uri=${encodedRedirectUri}` +
      `&response_mode=query` +
      `&scope=${encodedScope}` +
      `&state=${state}`;

    this.logger.log(`FINAL MICROSOFT LOGIN URL: ${authorizeUrl}`);

    return authorizeUrl;
  }

  /**
   * Exchange authorization code for tokens
   * @param code Authorization code
   * @param state Base64 encoded state string
   */
  async exchangeCodeForToken(code: string, state: string): Promise<TokenResponse> {
    // Parse the state
    const stateData = this.parseState(state);

    if (!stateData || !stateData.userId) {
      throw new Error('Invalid state parameter - missing user ID');
    }

    // Validate CSRF token (timestamp validation is now included in validateCsrfToken)
    const csrfError = await this.validateCsrfToken(stateData.csrf, stateData.timestamp);

    if (csrfError) {
      this.logger.error(`CSRF validation failed for user ${String(stateData.userId)}: ${csrfError}`);
      throw new Error(`CSRF validation failed: ${csrfError}`);
    }

    try {
      this.logger.log(`Exchanging code for token with redirect URI: ${this.redirectUri}`);

      const postData = new URLSearchParams({
        client_id: this.clientId,
        scope: this.scope,
        code: code,
        redirect_uri: this.redirectUri,
        grant_type: 'authorization_code',
        client_secret: this.clientSecret,
      });

      this.logger.log(`Token request payload: ${postData.toString()}`);

      const tokenResponse = await axios.post<MicrosoftTokenApiResponse>(
        this.tokenEndpoint,
        postData,
        {
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
          },
        },
      );

      // Convert the API response to our internal TokenResponse format
      const tokenData: TokenResponse = {
        access_token: tokenResponse.data.access_token,
        refresh_token: tokenResponse.data.refresh_token || '',
        expires_in: tokenResponse.data.expires_in,
      };

      // Emit event directly with parameters instead of a payload object
      await Promise.resolve(
        this.eventEmitter.emit(OutlookEventTypes.AUTH_TOKENS_SAVE, stateData.userId, tokenData),
      );

      // Create webhook subscription for the user's calendar
      try {
        await this.outlookService.createWebhookSubscription(
          Number(stateData.userId),
          tokenData.access_token,
          tokenData.refresh_token,
        );
        this.logger.log(`Successfully created webhook subscription for user ${String(stateData.userId)}`);
      } catch (webhookError) {
        // Don't fail authentication if webhook creation fails
        this.logger.error(
          `Failed to create webhook subscription: ${webhookError instanceof Error ? webhookError.message : 'Unknown error'}`,
        );
      }

      return tokenData;
    } catch (error) {
      this.logger.error(`Error exchanging code for token:`, error);
      throw new Error('Failed to exchange code for token');
    }
  }

  /**
   * Refresh an access token using a refresh token
   * @param refreshToken - The refresh token to use
   * @param userId - User ID associated with the token
   * @param calendarId - Calendar ID associated with the token
   * @returns New token response with refreshed access token and refresh token
   */
  async refreshAccessToken(
    refreshToken: string,
    userId?: number,
    calendarId?: string,
  ): Promise<TokenResponse> {
    try {
      const payload = {
        client_id: this.clientId,
        client_secret: this.clientSecret,
        refresh_token: refreshToken,
        grant_type: 'refresh_token',
        scope: this.scope,
      };

      const response = await axios.post<MicrosoftTokenApiResponse>(
        this.tokenEndpoint,
        qs.stringify(payload),
        {
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
          },
        },
      );

      // Validate required fields
      if (!response.data.access_token || !response.data.expires_in) {
        throw new Error('Invalid token refresh response from Microsoft');
      }

      // Microsoft might not return a new refresh token, in which case we should reuse the old one
      const newRefreshToken = response.data.refresh_token || refreshToken;

      const tokenData: TokenResponse = {
        access_token: response.data.access_token,
        refresh_token: newRefreshToken,
        expires_in: response.data.expires_in,
      };

      // If userId and calendarId are provided, emit event to update the token
      if (userId !== undefined && calendarId) {
        const userIdStr = String(userId);
        const calendarIdStr = String(calendarId);
        await Promise.resolve(
          this.eventEmitter.emit(
            OutlookEventTypes.AUTH_TOKENS_UPDATE,
            userIdStr,
            calendarIdStr,
            tokenData,
          ),
        );
      }

      return tokenData;
    } catch (error) {
      this.logger.error('Error refreshing access token:', error);
      throw new Error('Failed to refresh access token from Microsoft');
    }
  }

  /**
   * Renew a webhook subscription
   * @param subscriptionId - The ID of the subscription to renew
   * @param accessToken - The access token to use for renewal
   * @param refreshToken - The refresh token to use if access token needs refresh
   * @returns Updated webhook subscription with new expiration date
   */
  async renewWebhookSubscription(
    subscriptionId: string,
    accessToken: string,
    refreshToken: string,
  ): Promise<Subscription> {
    try {
      // Check if token needs refresh
      try {
        // First attempt with current access token
        return await this.outlookService.renewWebhookSubscription(subscriptionId, accessToken);
      } catch (error) {
        this.logger.warn(
          `Access token might be expired, attempting refresh: ${error instanceof Error ? error.message : 'Unknown error'}`,
        );

        // If token is expired, refresh it and try again
        const newTokens = await this.refreshAccessToken(refreshToken);
        return await this.outlookService.renewWebhookSubscription(
          subscriptionId,
          newTokens.access_token,
        );
      }
    } catch (error) {
      this.logger.error(
        `Failed to renew webhook subscription: ${error instanceof Error ? error.message : 'Unknown error'}`,
      );
      throw new Error('Failed to renew webhook subscription');
    }
  }

  /**
   * Check if a token is expired
   * @param tokenExpiry - Token expiry date
   * @param bufferMinutes - Buffer time in minutes before actual expiry to consider token expired
   * @returns Boolean indicating if token is expired
   */
  isTokenExpired(tokenExpiry: Date, bufferMinutes = 5): boolean {
    // Add buffer time to current time to prevent using tokens that will expire very soon
    const currentTimeWithBuffer = new Date(Date.now() + bufferMinutes * 60 * 1000);
    return tokenExpiry < currentTimeWithBuffer;
  }
}
