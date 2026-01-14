import { Injectable, Logger, Inject, forwardRef } from '@nestjs/common';
import { EventEmitter2 } from '@nestjs/event-emitter';
import axios from 'axios';
import { TokenResponse } from '../../interfaces/outlook/token-response.interface';
import { CalendarService } from '../calendar/calendar.service';
import { EmailService } from '../email/email.service';
import { MICROSOFT_CONFIG } from '../../constants';
import { MicrosoftOutlookConfig } from '../../interfaces/config/outlook-config.interface';
import { OutlookEventTypes } from '../../enums/event-types.enum';
import * as crypto from 'crypto';
import { Cron, CronExpression } from '@nestjs/schedule';
import { MicrosoftCsrfTokenRepository } from '../../repositories/microsoft-csrf-token.repository';
import { MicrosoftTokenApiResponse } from '../../interfaces/microsoft-auth/microsoft-token-api-response.interface';
import { StateObject } from '../../interfaces/microsoft-auth/state-object.interface';
import { PermissionScope } from '../../enums/permission-scope.enum';
import { InjectRepository } from '@nestjs/typeorm';
import { FindOptionsWhere, Repository } from 'typeorm';
import { MicrosoftUser } from '../../entities/microsoft-user.entity';
import { retryWithBackoff } from '../../utils/retry.util';

/**
 * Important terminology:
 *
 * - externalUserId (string): The ID of the user in the host application that uses this library.
 *   This is what we store in the MicrosoftUser entity to identify which external user
 *   the Microsoft tokens belong to.
 *
 * - internalUserId (number): The auto-generated primary key (MicrosoftUser.id) used for
 *   internal database relationships. Not exposed in public APIs.
 *
 * @see docs/USER_ID_TERMINOLOGY.md for detailed explanation
 */

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
  // Map to track subscription creation in progress for a user (keyed by external user ID)
  private subscriptionInProgress = new Map<string, boolean>();

  constructor(
    private readonly eventEmitter: EventEmitter2,
    @Inject(forwardRef(() => CalendarService))
    private readonly calendarService: CalendarService,
    @Inject(forwardRef(() => EmailService))
    private readonly emailService: EmailService,
    @Inject(MICROSOFT_CONFIG)
    private readonly microsoftConfig: MicrosoftOutlookConfig,
    private readonly csrfTokenRepository: MicrosoftCsrfTokenRepository,
    @InjectRepository(MicrosoftUser)
    private readonly microsoftUserRepository: Repository<MicrosoftUser>,
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
    const baseUrl = config.backendBaseUrl;

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
    }

    // Ensure the path doesn't have double slashes
    path = path.replace(/\/+/g, '/');

    const finalUri = `${cleanBaseUrl}${path}`;
    this.logger.debug(`Constructed redirect URI: ${finalUri}`);
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
   * @param externalUserId - The external user ID from the host application
   */
  private async generateCsrfToken(externalUserId: string | number): Promise<string> {
    const token = crypto.randomBytes(32).toString('hex');

    // Save token in the database
    await this.csrfTokenRepository.saveToken(token, externalUserId, this.CSRF_TOKEN_EXPIRY);

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
   * @param externalUserId - External user ID from the host application
   * @param scopes - Optional array of permission scopes, uses default scopes if not provided
   */
  async getLoginUrl(
    externalUserId: string,
    scopes: PermissionScope[] = this.defaultScopes
  ): Promise<string> {
    // Generate a secure CSRF token linked to this user
    const csrf = await this.generateCsrfToken(externalUserId);

    // Generate state with user ID, CSRF token, and requested scopes
    const stateObj = {
      userId: externalUserId,
      csrf,
      timestamp: Date.now(),
      requestedScopes: scopes,
    };
    const stateJson = JSON.stringify(stateObj);
    const state = Buffer.from(stateJson).toString('base64').replace(/=/g, ''); // Remove padding '=' characters

    this.logger.debug(`State object: ${JSON.stringify(stateObj)}`);

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

    this.logger.debug(`Final Microsoft login URL: ${authorizeUrl}`);

    return authorizeUrl;
  }

  /**
   * Save a Microsoft user with token information and scopes
   * On reconnection, reuses existing inactive user record instead of creating duplicates
   */
  private async saveMicrosoftUser(
    externalUserId: string,
    accessToken: string,
    refreshToken: string,
    expiresIn: number,
    scopes: string
  ): Promise<void> {
    // Find existing user (including inactive ones) or create a new one
    let user = await this.microsoftUserRepository.findOne({
      where: { externalUserId: externalUserId }
    });

    if (!user) {
      user = new MicrosoftUser();
      user.externalUserId = externalUserId;
      this.logger.log(`Creating new Microsoft user for external user ${externalUserId}`);
    } else {
      this.logger.log(`Reusing existing Microsoft user record (id: ${user.id}) for external user ${externalUserId}`);
    }

    // Update token information
    user.accessToken = accessToken;
    user.refreshToken = refreshToken;
    user.tokenExpiry = new Date(Date.now() + expiresIn * 1000);
    user.scopes = scopes;
    user.isActive = true; // Reactivate if previously inactive

    await this.microsoftUserRepository.save(user);
  }

  /**
   * Get Microsoft user by internal ID or external user ID
   *
   * @param params - Object with either internalUserId or externalUserId
   * @returns Microsoft user entity or null if not found
   */
  private async getMicrosoftUser(params: { internalUserId?: number; externalUserId?: string }): Promise<MicrosoftUser | null> {
    const { internalUserId, externalUserId } = params;

    if (!internalUserId && !externalUserId) {
      throw new Error('Either internalUserId or externalUserId must be provided');
    }

    const whereCondition: FindOptionsWhere<MicrosoftUser> = { isActive: true };

    if (internalUserId !== undefined) {
      whereCondition.id = internalUserId;
    }

    if (externalUserId !== undefined) {
      whereCondition.externalUserId = externalUserId;
    }

    return await this.microsoftUserRepository.findOne({
      where: whereCondition
    });
  }

  /**
   * Gets a valid access token for a user, REFRESH it if necessary
   *
   * @param params - Object with either internalUserId or externalUserId
   * @returns Valid access token string
   */
  async getUserAccessToken(params: { internalUserId?: number; externalUserId?: string }): Promise<string> {
    try {
      const { internalUserId, externalUserId } = params;

      // Guard clause
      if (!internalUserId && !externalUserId) {
        throw new Error('Either internalUserId or externalUserId must be provided');
      }

      // Get the Microsoft user by internal or external user ID
      const user = await this.getMicrosoftUser({
        internalUserId,
        externalUserId
      });

      // Guard clause
      if (!user) {
        const identifier = internalUserId ? `internal ID ${String(internalUserId)}` : `external ID ${externalUserId}`;
        throw new Error(`No Microsoft user found with ${identifier}`);
      }

      // Extract & REFRESH the token information
      return await this.processTokenInfo({
        accessToken: user.accessToken,
        refreshToken: user.refreshToken,
        tokenExpiry: user.tokenExpiry,
        scopes: user.scopes
      }, user.id);
    } 
    
    // Error handling
    catch (error) {
      const identifier = params.internalUserId
        ? `internal user ID ${String(params.internalUserId)}`
        : `external user ID ${params.externalUserId}`;
      this.logger.error(`Error getting access token for ${identifier}:`, error);
      throw new Error(`Failed to get valid access token: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  /**
   * Common helper to process token information and refresh if needed
   * @param tokenInfo - Token information with access token, refresh token, expiry date and scopes
   * @param userId - Internal user ID for refreshing tokens
   * @returns Valid access token
   */
  private async processTokenInfo(
    tokenInfo: {
      accessToken: string;
      refreshToken: string;
      tokenExpiry: Date;
      scopes: string;
    },
    internalUserId: number
  ): Promise<string> {
    // Check if the token is still valid
    if (!this.isTokenExpired(tokenInfo.tokenExpiry)) {
      // Token is still valid, return it
      return tokenInfo.accessToken;
    }
    
    // Token is expired, refresh it
    this.logger.log(`Access token for user ID ${String(internalUserId)} is expired, refreshing...`);
    
    const accessToken = await this.refreshAccessToken(tokenInfo.refreshToken, internalUserId);
    
    return accessToken;
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

    const correlationId = `auth-${stateData.userId}-${Date.now()}`;
    this.logger.log(`[${correlationId}] Starting token exchange for user ${stateData.userId}`);

    // Validate CSRF token (timestamp validation is now included in validateCsrfToken)
    const csrfError = await this.validateCsrfToken(stateData.csrf, stateData.timestamp);

    if (csrfError) {
      this.logger.error(`[${correlationId}] CSRF validation failed for user ${String(stateData.userId)}: ${csrfError}`);
      throw new Error(`CSRF validation failed: ${csrfError}`);
    }

    try {
      this.logger.log(`[${correlationId}] Exchanging code for token with redirect URI: ${this.redirectUri}`);

      // Use scopes from state if available, otherwise use defaults
      const scopesToUse = stateData.requestedScopes || this.defaultScopes;
      this.logger.log(`[${correlationId}] Using scopes for token exchange: ${scopesToUse.join(', ')}`);

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

      this.logger.debug(`[${correlationId}] Token request payload: ${postData.toString()}`);

      const tokenResponse = await axios.post<MicrosoftTokenApiResponse>(
        this.tokenEndpoint,
        postData,
        {
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
          },
        },
      );

      this.logger.log(`[${correlationId}] Successfully received token from Microsoft`);

      // Convert the API response to our internal TokenResponse format
      const tokenData: TokenResponse = {
        access_token: tokenResponse.data.access_token,
        refresh_token: tokenResponse.data.refresh_token || '',
        expires_in: tokenResponse.data.expires_in,
      };

      // Save Microsoft user with their tokens and scopes for later use
      this.logger.log(`[${correlationId}] Saving Microsoft user to database`);
      await this.saveMicrosoftUser(
        stateData.userId,
        tokenData.access_token,
        tokenData.refresh_token,
        tokenData.expires_in,
        scopeString // Store the exact Microsoft scopes used
      );

      // Emit event that the user has been authenticated
      this.logger.log(`[${correlationId}] Emitting USER_AUTHENTICATED event`);
      await Promise.resolve(
        this.eventEmitter.emit(OutlookEventTypes.USER_AUTHENTICATED, stateData.userId, {
          externalUserId: stateData.userId,
          scopes: scopesToUse
        }),
      );

      // Setup subscriptions (both calendar and email)
      this.logger.log(`[${correlationId}] Starting subscription setup (async)`);
      await this.setupSubscriptions(stateData.userId, scopesToUse);

      this.logger.log(`[${correlationId}] Token exchange completed successfully`);
      return tokenData;
    } catch (error) {
      this.logger.error(`[${correlationId}] Error exchanging code for token:`, error);
      throw new Error('Failed to exchange code for token');
    }
  }
  
  /**
   * Setup webhook subscriptions for a user based on requested scopes
   * @param externalUserId - External user ID from the host application
   * @param scopes - Requested permission scopes
   */
  private async setupSubscriptions(
    externalUserId: string,
    scopes: PermissionScope[] = this.defaultScopes
  ): Promise<void> {
    // Check if subscription setup is already in progress for this user
    if (this.subscriptionInProgress.get(externalUserId)) {
      this.logger.log(`Subscription setup already in progress for user ${externalUserId}`);
      return;
    }

    const correlationId = `auth-${externalUserId}-${Date.now()}`;
    this.logger.log(`[${correlationId}] Starting subscription setup for user ${externalUserId}`);

    try {
      // Mark subscription setup as in progress
      this.subscriptionInProgress.set(externalUserId, true);

      const errors: string[] = [];

      // Check if calendar.read permissions were requested
      if (this.hasCalendarSubscriptionPermission(scopes)) {
        // Create webhook subscription for the user's calendar with retry
        try {
          this.logger.log(`[${correlationId}] Creating calendar webhook subscription with retry logic`);
          await retryWithBackoff(
            () => this.calendarService.createWebhookSubscription(externalUserId),
            {
              maxRetries: 3,
              retryDelayMs: 1000
            }
          );
          this.logger.log(`[${correlationId}] Successfully created calendar webhook subscription for user ${externalUserId}`);
        } catch (calendarError) {
          // Don't fail authentication if webhook creation fails
          const errorMsg = `Failed to create calendar webhook subscription after retries: ${calendarError instanceof Error ? calendarError.message : 'Unknown error'}`;
          this.logger.error(`[${correlationId}] ${errorMsg}`);
          errors.push(errorMsg);
        }
      }

      // Check if email.read permissions were requested
      if (this.hasEmailSubscriptionPermission(scopes)) {
        // Create webhook subscription for the user's email with retry
        try {
          this.logger.log(`[${correlationId}] Creating email webhook subscription with retry logic`);
          await retryWithBackoff(
            () => this.emailService.createWebhookSubscription(externalUserId),
            {
              maxRetries: 3,
              retryDelayMs: 1000
            }
          );
          this.logger.log(`[${correlationId}] Successfully created email webhook subscription for user ${externalUserId}`);
        } catch (emailError) {
          // Don't fail authentication if webhook creation fails
          const errorMsg = `Failed to create email webhook subscription after retries: ${emailError instanceof Error ? emailError.message : 'Unknown error'}`;
          this.logger.error(`[${correlationId}] ${errorMsg}`);
          errors.push(errorMsg);
        }
      }

      if (errors.length > 0) {
        this.logger.error(
          `[${correlationId}] Subscription setup completed with errors for user ${externalUserId}: ${errors.join('; ')}`
        );
      } else {
        this.logger.log(`[${correlationId}] All subscriptions created successfully for user ${externalUserId}`);
      }
    } catch (error) {
      this.logger.error(
        `[${correlationId}] Unexpected error setting up subscriptions: ${error instanceof Error ? error.message : 'Unknown error'}`
      );
      // Continue without failing authentication
    } finally {
      // Mark subscription setup as complete
      this.subscriptionInProgress.set(externalUserId, false);
      this.logger.log(`[${correlationId}] Subscription setup process finished for user ${externalUserId}`);
    }
  }

  /**
   * Refresh an access token using a refresh token
   *
   * ⚠️ **ADVANCED USAGE**: This method uses internal database IDs (MicrosoftUser.id).
   *
   * **Most applications should use `getUserAccessToken({ internalUserId })` instead**,
   * which handles token refresh automatically without needing to manage tokens manually.
   *
   * **Only use this method if:**
   * - You're directly accessing and managing `MicrosoftUser` entities
   * - You need fine-grained control over the token refresh process
   * - You already have the refresh token and internal user ID from the database
   *
   * **Typical usage (advanced):**
   * ```typescript
   * const microsoftUser = await microsoftUserRepo.findOne({...});
   * if (isTokenExpired(microsoftUser.tokenExpiry)) {
   *   const newToken = await authService.refreshAccessToken(
   *     microsoftUser.refreshToken,
   *     microsoftUser.id  // Internal database ID
   *   );
   * }
   * ```
   *
   * @param refreshToken - The refresh token to use for obtaining a new access token
   * @param internalUserId - Internal database user ID (MicrosoftUser.id primary key)
   * @returns New access token
   *
   * @see getUserAccessToken For standard usage with automatic token refresh
   */
  async refreshAccessToken(
    refreshToken: string,
    internalUserId: number
  ): Promise<string> {
    try {
      // Get the user to access its properties
      const internalUser = await this.microsoftUserRepository.findOne({
        where: { id: internalUserId }
      });

      if (!internalUser) {
        throw new Error(`No user found with ID ${String(internalUserId)}`);
      }
      
      const scopeString = internalUser.scopes;
      this.logger.debug(`Using saved scopes from database: ${scopeString}`);

      this.logger.debug(`Refreshing token for user ID ${String(internalUserId)} with scopes: ${scopeString}`);
      
      // Prepare parameters as specified in Microsoft documentation
      const payload = new URLSearchParams({
        client_id: this.clientId,
        client_secret: this.clientSecret,
        refresh_token: refreshToken,
        grant_type: 'refresh_token',
        scope: scopeString,
      });
      
      try {
        const response = await axios.post<MicrosoftTokenApiResponse>(
          this.tokenEndpoint,
          payload.toString(),
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
        const newAccessToken = response.data.access_token;

        // Update Microsoft user record with new tokens
        internalUser.accessToken = newAccessToken;
        internalUser.refreshToken = newRefreshToken;
        internalUser.tokenExpiry = new Date(Date.now() + response.data.expires_in * 1000);
        
        await this.microsoftUserRepository.save(internalUser);

        // Return just the access token
        return newAccessToken;
      } catch (error) {
        if (axios.isAxiosError(error) && error.response) {
          // Log detailed API error information
          this.logger.error(
            `Microsoft API error refreshing token for user ID ${String(internalUserId)}: Status: ${String(error.response.status)}, Response: ${JSON.stringify(error.response.data)}`
          );
          
          // Check for specific error conditions from Microsoft
          const errorData = error.response.data as { error?: string };
          if (errorData.error === 'invalid_grant') {
            throw new Error('Microsoft refresh token is invalid or expired');
          }
        }
        throw error; // Re-throw for the outer catch to handle
      }
    } catch (error) {
      this.logger.error(`Error refreshing access token for user ID ${String(internalUserId)}:`, error);
      throw new Error('Failed to refresh access token from Microsoft');
    }
  }

  /**
   * Revoke Microsoft tokens using the refresh token
   * @param refreshToken - The refresh token to use
   * @returns void
   */
  async revokeRefreshToken(refreshToken: string): Promise<void> {
    try {
      if (!refreshToken) {
        this.logger.warn('⚠️ No refresh token available for revocation');
        return;
      }

      await axios.post(
        'https://login.microsoftonline.com/common/oauth2/v2.0/logout',
        new URLSearchParams({
          token: refreshToken,
          token_type_hint: 'refresh_token',
        }),
        {
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        },
      );

      this.logger.log('✅ Microsoft tokens revoked successfully');
    } catch (error) {
      this.logger.warn(
        `⚠️ Failed to revoke Microsoft tokens: ${error instanceof Error ? error.message : 'Unknown error'}`,
      );
    }
  }

  /**
   * Helper method to determine if calendar permissions were requested
   */
  private hasCalendarSubscriptionPermission(scopes: PermissionScope[]): boolean {
    return scopes.some(scope => 
      scope === PermissionScope.CALENDAR_READ
    );
  }

  /**
   * Helper method to determine if email permissions were requested
   */  
  private hasEmailSubscriptionPermission(scopes: PermissionScope[]): boolean {
    return scopes.some(scope => 
      scope === PermissionScope.EMAIL_READ
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