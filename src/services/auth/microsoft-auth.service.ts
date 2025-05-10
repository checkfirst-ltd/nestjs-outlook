import { Injectable, Logger, Inject, forwardRef } from '@nestjs/common';
import { EventEmitter2 } from '@nestjs/event-emitter';
import axios from 'axios';
import { TokenResponse } from '../../interfaces/outlook/token-response.interface';
import * as qs from 'querystring';
import { CalendarService } from '../calendar/calendar.service';
import { EmailService } from '../email/email.service';
import { Subscription } from '@microsoft/microsoft-graph-types';
import { MICROSOFT_CONFIG } from '../../constants';
import { MicrosoftOutlookConfig } from '../../interfaces/config/outlook-config.interface';
import { OutlookEventTypes } from '../../enums/event-types.enum';
import * as crypto from 'crypto';
import { Cron, CronExpression } from '@nestjs/schedule';
import { MicrosoftCsrfTokenRepository } from '../../repositories/microsoft-csrf-token.repository';
import { MicrosoftTokenApiResponse } from '../../interfaces/microsoft-auth/microsoft-token-api-response.interface';
import { StateObject } from '../../interfaces/microsoft-auth/state-object.interface';
import { PermissionScope } from '../../enums/permission-scope.enum';

@Injectable()
export class MicrosoftAuthService {
  private readonly logger = new Logger(MicrosoftAuthService.name);
  private readonly clientId: string;
  private readonly clientSecret: string;
  private readonly tenantId = 'common';
  private readonly redirectUri: string;
  private readonly tokenEndpoint: string;
  // Required Microsoft scopes that are always included
  private readonly requiredScopes = ['offline_access', 'User.Read'];
  private readonly defaultScopes: PermissionScope[] = [
    PermissionScope.CALENDAR_READ,
    PermissionScope.CALENDAR_WRITE,
    PermissionScope.EMAIL_SEND,
    PermissionScope.EMAIL_READ,
    PermissionScope.EMAIL_WRITE,
  ];
  // CSRF token expiration time (30 minutes)
  private readonly CSRF_TOKEN_EXPIRY = 30 * 60 * 1000;
  // Map to track subscription creation in progress for a user
  private subscriptionInProgress = new Map<number, boolean>();

  constructor(
    private readonly eventEmitter: EventEmitter2,
    @Inject(forwardRef(() => CalendarService))
    private readonly calendarService: CalendarService,
    @Inject(forwardRef(() => EmailService))
    private readonly emailService: EmailService,
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

    // Log the redirect URI to help with debugging
    this.logger.log(`Microsoft OAuth redirect URI set to: ${this.redirectUri}`);
  }

  /**
   * Maps generic permission scopes to Microsoft-specific permission scopes
   */
  private mapToMicrosoftScopes(scopes: PermissionScope[]): string[] {
    const scopeMapping: Record<PermissionScope, string[]> = {
      [PermissionScope.CALENDAR_READ]: ['Calendars.Read'],
      [PermissionScope.CALENDAR_WRITE]: ['Calendars.ReadWrite'],
      [PermissionScope.EMAIL_READ]: ['Mail.Read'],
      [PermissionScope.EMAIL_WRITE]: ['Mail.ReadWrite'],
      [PermissionScope.EMAIL_SEND]: ['Mail.Send'],
    };

    // Flatten and deduplicate scopes
    const microsoftScopes = new Set<string>();
    
    // Add required scopes
    this.requiredScopes.forEach(scope => microsoftScopes.add(scope));
    
    // Add mapped scopes
    scopes.forEach(scope => {
      scopeMapping[scope].forEach(mappedScope => microsoftScopes.add(mappedScope));
    });
    
    return Array.from(microsoftScopes);
  }

  /**
   * Builds redirect URI based on configuration values
   * Format: backendBaseUrl/[basePath]/redirectPath
   * @param config Microsoft Outlook Configuration
   * @returns Complete redirect URI
   */
  private buildRedirectUri(config: MicrosoftOutlookConfig): string {
    // If redirectPath already contains a full URL, use it directly
    if (config.redirectPath.startsWith('http')) {
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
   * @param userId User ID
   * @param scopes Optional array of permission scopes, uses default scopes if not provided
   */
  async getLoginUrl(
    userId: string, 
    scopes: PermissionScope[] = this.defaultScopes
  ): Promise<string> {
    // Generate a secure CSRF token linked to this user
    const csrf = await this.generateCsrfToken(userId);

    // Generate state with user ID, CSRF token, and requested scopes
    const stateObj = {
      userId,
      csrf,
      timestamp: Date.now(),
      requestedScopes: scopes,
    };
    const stateJson = JSON.stringify(stateObj);
    const state = Buffer.from(stateJson).toString('base64').replace(/=/g, ''); // Remove padding '=' characters

    this.logger.log(`State object: ${JSON.stringify(stateObj)}`);

    // Build scope string and encode it
    const scopeString = this.mapToMicrosoftScopes(scopes).join(' ');
    const encodedScope = encodeURIComponent(scopeString);
    
    // Ensure proper URI encoding for parameters
    const encodedRedirectUri = encodeURIComponent(this.redirectUri);

    this.logger.debug(`Requested generic scopes: ${scopes.join(', ')}`);
    this.logger.debug(`Mapped to Microsoft scopes: ${scopeString}`);
    this.logger.debug(`Redirect URI (raw): ${this.redirectUri}`);
    this.logger.debug(`Redirect URI (encoded): ${encodedRedirectUri}`);

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
  async exchangeCodeForToken(
    code: string, 
    state: string
  ): Promise<TokenResponse> {
    // Parse the state
    const stateData = this.parseState(state);

    if (!stateData?.userId) {
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
      
      // Use scopes from state if available, otherwise use defaults
      const scopesToUse = stateData.requestedScopes || this.defaultScopes;
      this.logger.log(`Using scopes for token exchange: ${scopesToUse.join(', ')}`);
      
      // Build scope string
      const scopeString = this.mapToMicrosoftScopes(scopesToUse).join(' ');

      const postData = new URLSearchParams({
        client_id: this.clientId,
        scope: scopeString,
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

      const userId = Number(stateData.userId);
      const accessToken = tokenData.access_token;
      const refreshToken = tokenData.refresh_token;

      // Setup subscriptions (both calendar and email)
      await this.setupSubscriptions(userId, accessToken, refreshToken, scopesToUse);

      return tokenData;
    } catch (error) {
      this.logger.error(`Error exchanging code for token:`, error);
      throw new Error('Failed to exchange code for token');
    }
  }

  /**
   * Setup webhook subscriptions for a user based on requested scopes
   * @param userId User ID
   * @param accessToken Access token
   * @param refreshToken Refresh token
   * @param scopes Requested permission scopes
   */
  private async setupSubscriptions(
    userId: number, 
    accessToken: string, 
    refreshToken: string,
    scopes: PermissionScope[] = this.defaultScopes
  ): Promise<void> {
    // Check if subscription setup is already in progress for this user
    if (this.subscriptionInProgress.get(userId)) {
      this.logger.log(`Subscription setup already in progress for user ${String(userId)}`);
      return;
    }

    try {
      // Mark subscription setup as in progress
      this.subscriptionInProgress.set(userId, true);

      // Check if calendar permissions were requested
      if (this.hasCalendarPermission(scopes)) {
        // Create webhook subscription for the user's calendar
        try {
          await this.calendarService.createWebhookSubscription(
            userId,
            accessToken,
            refreshToken,
          );
          this.logger.log(`Successfully created calendar webhook subscription for user ${String(userId)}`);
        } catch (calendarError) {
          // Don't fail authentication if webhook creation fails
          this.logger.error(
            `Failed to create calendar webhook subscription: ${calendarError instanceof Error ? calendarError.message : 'Unknown error'}`,
          );
        }
      }

      // Check if email permissions were requested
      if (this.hasEmailPermission(scopes)) {
        // Create webhook subscription for the user's email
        try {
          await this.emailService.createWebhookSubscription(
            userId,
            accessToken,
            refreshToken,
          );
          this.logger.log(`Successfully created email webhook subscription for user ${String(userId)}`);
        } catch (emailError) {
          // Don't fail authentication if webhook creation fails
          this.logger.error(
            `Failed to create email webhook subscription: ${emailError instanceof Error ? emailError.message : 'Unknown error'}`,
          );
        }
      }
    } catch (error) {
      this.logger.error(`Error setting up subscriptions: ${error instanceof Error ? error.message : 'Unknown error'}`);
      // Continue without failing authentication
    } finally {
      // Mark subscription setup as complete
      this.subscriptionInProgress.set(userId, false);
    }
  }

  /**
   * Refresh an access token using a refresh token
   * @param refreshToken - The refresh token to use
   * @param userId - User ID associated with the token
   * @param calendarId - Calendar ID associated with the token
   * @param scopes - Permission scopes to request in the refresh
   * @returns New token response with refreshed access token and refresh token
   */
  async refreshAccessToken(
    refreshToken: string,
    userId?: number,
    calendarId?: string,
    scopes: PermissionScope[] = this.defaultScopes
  ): Promise<TokenResponse> {
    try {
      const scopeString = this.mapToMicrosoftScopes(scopes).join(' ');
      
      const payload = {
        client_id: this.clientId,
        client_secret: this.clientSecret,
        refresh_token: refreshToken,
        grant_type: 'refresh_token',
        scope: scopeString,
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
   * Renew webhook subscription with refreshed token if needed
   */
  async renewWebhookSubscription(
    subscriptionId: string,
    accessToken: string,
    refreshToken: string,
    scopes: PermissionScope[] = this.defaultScopes
  ): Promise<Subscription> {
    try {
      // Try to renew with current token
      return await this.calendarService.renewWebhookSubscription(subscriptionId, accessToken);
    } catch (error: unknown) {
      if (axios.isAxiosError(error) && error.response?.status === 401) {
        this.logger.log('Access token expired during webhook renewal, refreshing token...');
        
        // Refresh the token
        const tokenResponse = await this.refreshAccessToken(refreshToken, undefined, undefined, scopes);
        
        // Retry with the new token
        return await this.calendarService.renewWebhookSubscription(
          subscriptionId,
          tokenResponse.access_token,
        );
      }
      
      throw error;
    }
  }

  /**
   * Helper method to determine if calendar permissions were requested
   */
  private hasCalendarPermission(scopes: PermissionScope[]): boolean {
    return scopes.some(scope => 
      scope === PermissionScope.CALENDAR_READ || 
      scope === PermissionScope.CALENDAR_WRITE
    );
  }

  /**
   * Helper method to determine if email permissions were requested
   */  
  private hasEmailPermission(scopes: PermissionScope[]): boolean {
    return scopes.some(scope => 
      scope === PermissionScope.EMAIL_READ || 
      scope === PermissionScope.EMAIL_WRITE || 
      scope === PermissionScope.EMAIL_SEND
    );
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