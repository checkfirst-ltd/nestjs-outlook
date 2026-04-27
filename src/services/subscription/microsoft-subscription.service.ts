import { randomUUID } from "crypto";
import { Injectable, Logger, Inject, forwardRef } from '@nestjs/common';
import { EventEmitter2 } from '@nestjs/event-emitter';
import { Cron } from '@nestjs/schedule';
import { InjectRepository } from '@nestjs/typeorm';
import { Repository, Not, In } from 'typeorm';
import axios from 'axios';
import { Subscription, BatchRequestPayload, BatchResponsePayload, BatchResponse } from '../../types';
import { MicrosoftAuthService } from '../auth/microsoft-auth.service';
import { OutlookWebhookSubscriptionRepository } from '../../repositories/outlook-webhook-subscription.repository';
import { OutlookWebhookSubscription } from '../../entities/outlook-webhook-subscription.entity';
import { MicrosoftUser } from '../../entities/microsoft-user.entity';
import { MICROSOFT_CONFIG } from '../../constants';
import { MicrosoftOutlookConfig } from '../../interfaces/config/outlook-config.interface';
import { OutlookEventTypes } from '../../enums/event-types.enum';
import { MicrosoftUserStatus } from '../../enums/microsoft-user-status.enum';
import { UserIdConverterService } from '../shared/user-id-converter.service';
import { GraphRateLimiterService } from '../shared/graph-rate-limiter.service';
import { executeGraphApiCall } from '../../utils/outlook-api-executor.util';

/**
 * Microsoft Graph API subscription structure
 */
export interface MicrosoftSubscription {
  id: string;
  resource: string;
  changeType?: string;
  clientState?: string;
  notificationUrl?: string;
  expirationDateTime?: string;
  creatorId?: string;
}

/**
 * Filter function for subscriptions
 */
export type SubscriptionFilter = (subscription: MicrosoftSubscription) => boolean;

/**
 * Options for subscription cleanup
 */
export interface SubscriptionCleanupOptions {
  accessToken: string;
  filter?: SubscriptionFilter;
}

/**
 * Result of subscription cleanup operation
 */
export interface SubscriptionCleanupResult {
  totalFound: number;
  successfullyDeleted: number;
  failedToDelete: number;
  deletedSubscriptionIds: string[];
  errors: Array<{ subscriptionId: string; error: string }>;
}

// ─── Health-check tuning constants ──────────────────────────────────
/** Microsoft Graph batch API allows at most 20 inner requests per /$batch call. */
const GRAPH_BATCH_SIZE = 20;
/** Only subs expiring within this window get a Graph /subscriptions/{id} verification call. */
const HEALTH_CHECK_WINDOW_HOURS = 24;
/** A sub expiring within this window is force-renewed during the health check. */
const FORCE_RENEWAL_WINDOW_HOURS = 12;
/** A sub with no notifications in this window is treated as stale. */
const STALE_NOTIFICATION_THRESHOLD_HOURS = 24;

/**
 * Per-pass counters for the health check. Helpers return a freshly-allocated
 * instance and the orchestrator sums them — no shared mutable state.
 */
interface HealthCheckCounters {
  verified: number;
  recreated: number;
  renewed: number;
  staleDetected: number;
  failed: number;
}

/**
 * Centralized service for managing Microsoft Graph API subscriptions
 * Handles creation, renewal, deletion, health checks, and cleanup of
 * webhook subscriptions for calendar, email, and other Microsoft integrations
 */
@Injectable()
export class MicrosoftSubscriptionService {
  private readonly logger = new Logger(MicrosoftSubscriptionService.name);
  private readonly graphApiBaseUrl = 'https://graph.microsoft.com/v1.0';
  private readonly msAuthUrl = 'https://login.microsoftonline.com/common/oauth2/v2.0';

  constructor(
    @Inject(forwardRef(() => MicrosoftAuthService))
    private readonly microsoftAuthService: MicrosoftAuthService,
    private readonly webhookSubscriptionRepository: OutlookWebhookSubscriptionRepository,
    private readonly eventEmitter: EventEmitter2,
    @Inject(MICROSOFT_CONFIG)
    private readonly microsoftConfig: MicrosoftOutlookConfig,
    @InjectRepository(MicrosoftUser)
    private readonly microsoftUserRepository: Repository<MicrosoftUser>,
    private readonly userIdConverter: UserIdConverterService,
    private readonly rateLimiter: GraphRateLimiterService,
  ) {}

  // ─── Microsoft Graph API subscription queries ───────────────────────

  /**
   * Get all active subscriptions from Microsoft Graph API
   * @param accessToken - Valid Microsoft access token
   * @returns Array of all subscriptions
   */
  async getActiveSubscriptions(accessToken: string): Promise<MicrosoftSubscription[]> {
    try {
      const response = await executeGraphApiCall(
        () => axios.get(`${this.graphApiBaseUrl}/subscriptions`, {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/json',
            'Prefer': 'IdType="ImmutableId"',
          },
          timeout: 10000,
        }),
        {
          logger: this.logger,
          resourceName: 'get active subscriptions',
          maxRetries: 7,
        }
      );

      const data = response?.data as { value?: MicrosoftSubscription[] };

      return data.value || [];
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      this.logger.warn(`Failed to get subscriptions from Microsoft: ${message}`);
      return [];
    }
  }

  /**
   * Get subscriptions filtered by client ID pattern (client_{id}_)
   * @param clientId - The client ID to filter by
   * @param accessToken - Valid Microsoft access token
   * @returns Array of matching subscriptions
   */
  async getActiveSubscriptionsForClient(
    clientId: number,
    accessToken: string,
  ): Promise<MicrosoftSubscription[]> {
    const allSubscriptions = await this.getActiveSubscriptions(accessToken);
    return allSubscriptions.filter((sub) => sub.clientState?.includes(`client_${clientId}_`));
  }

  /**
   * Get subscriptions filtered by user ID pattern (user_{id}_)
   * @param userId - The user ID to filter by
   * @param accessToken - Valid Microsoft access token
   * @returns Array of matching subscriptions
   */
  async getActiveSubscriptionsForUser(
    userId: number,
    accessToken: string,
  ): Promise<MicrosoftSubscription[]> {
    const allSubscriptions = await this.getActiveSubscriptions(accessToken);
    return allSubscriptions.filter((sub) => sub.clientState?.includes(`user_${userId}_`));
  }

  // ─── Single subscription CRUD ───────────────────────────────────────

  /**
   * Delete a single subscription from Microsoft Graph API
   * @param subscriptionId - The subscription ID to delete
   * @param accessToken - Valid Microsoft access token
   * @throws Error if deletion fails (except 404 which is handled gracefully)
   */
  async deleteSubscription(subscriptionId: string, accessToken: string): Promise<void> {
    try {
      await executeGraphApiCall(
        () => axios.delete(`${this.graphApiBaseUrl}/subscriptions/${subscriptionId}`, {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            'Prefer': 'IdType="ImmutableId"',
          },
          timeout: 10000,
        }),
        {
          logger: this.logger,
          resourceName: `delete subscription ${subscriptionId}`,
          maxRetries: 7,
          return404AsNull: true,
        }
      );
      this.logger.log(`✅ Deleted subscription ${subscriptionId} at Microsoft`);
    } catch (error: unknown) {
      throw error;
    }
  }

  // ─── Webhook subscription lifecycle ─────────────────────────────────

  /**
   * Creates a webhook subscription to receive notifications for calendar events
   * @param externalUserId - External user ID
   * @returns The created subscription data
   */
  async createWebhookSubscription(
    externalUserId: string,
  ): Promise<Subscription> {
    // Convert external user ID to internal database ID
    const internalUserId = await this.userIdConverter.externalToInternal(externalUserId, {cache: false});

    const correlationId = `webhook-${internalUserId}-${Date.now()}`;
    this.logger.log(`[${correlationId}] Starting webhook subscription creation for user ${internalUserId}`);

    try {
      // Get a valid access token for this user
      this.logger.log(`[${correlationId}] Fetching access token for user ${internalUserId}`);

      const accessToken =
        await this.microsoftAuthService.getUserAccessToken({internalUserId, cache: false});

      this.logger.log(`[${correlationId}] Successfully obtained access token`);

      // Set expiration date (max 3 days as per Microsoft documentation)
      const expirationDateTime = new Date();
      expirationDateTime.setHours(expirationDateTime.getHours() + 72); // 3 days from now

      const appUrl =
        this.microsoftConfig.backendBaseUrl || "http://localhost:3000";
      const basePath = this.microsoftConfig.basePath;
      const basePathUrl = basePath ? `${appUrl}/${basePath}` : appUrl;

      const webhookPath = this.microsoftConfig.calendarWebhookPath || '/calendar/webhook';
      const notificationUrl = `${basePathUrl}${webhookPath}`;

      // Create subscription payload
      const subscriptionData = {
        changeType: "created,updated,deleted",
        notificationUrl,
        // Add lifecycleNotificationUrl for increased reliability
        lifecycleNotificationUrl: notificationUrl,
        resource: "/me/events",
        expirationDateTime: expirationDateTime.toISOString(),
        clientState: `user_${internalUserId}_${randomUUID()}`,
      };

      this.logger.log(
        `[${correlationId}] Creating webhook subscription with notificationUrl: ${notificationUrl}`
      );

      this.logger.debug(
        `[${correlationId}] Subscription data: ${JSON.stringify(subscriptionData)}`
      );
      // Create the subscription with Microsoft Graph API
      this.logger.log(`[${correlationId}] Sending POST request to Microsoft Graph API`);
      const response = await executeGraphApiCall(
        () => axios.post<Subscription>(
          `${this.graphApiBaseUrl}/subscriptions`,
          subscriptionData,
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
              "Content-Type": "application/json",
              "Prefer": 'IdType="ImmutableId"',
            },
          }
        ),
        {
          logger: this.logger,
          resourceName: `create webhook subscription for user ${internalUserId}`,
          maxRetries: 7,
        }
      );

      if (!response?.data) {
        throw new Error('Subscription creation returned null response');
      }

      this.logger.log(
        `[${correlationId}] Created webhook subscription ${response.data.id || "unknown"} for user ${internalUserId}`
      );

      // Save the subscription to the database
      this.logger.log(`[${correlationId}] Saving subscription to database (internalUserId: ${internalUserId}, externalUserId: ${externalUserId})`);
      await this.webhookSubscriptionRepository.saveSubscription({
        subscriptionId: response.data.id,
        userId: internalUserId,
        resource: response.data.resource,
        changeType: response.data.changeType,
        clientState: response.data.clientState || "",
        notificationUrl: response.data.notificationUrl,
        expirationDateTime: response.data.expirationDateTime
          ? new Date(response.data.expirationDateTime)
          : new Date(),
      });

      this.logger.log(`[${correlationId}] Successfully stored subscription in database`);
      this.logger.log(`[${correlationId}] Webhook subscription creation completed successfully`);

      return response.data;
    } catch (error) {
      if (axios.isAxiosError(error)) {
        this.logger.error(
          `[${correlationId}] Microsoft Graph API error: Status ${error.response?.status}, ` +
          `Message: ${JSON.stringify(error.response?.data)}`
        );
      } else {
        this.logger.error(`[${correlationId}] Failed to create webhook subscription:`, error);
      }
      throw new Error(`Failed to create webhook subscription: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  /**
   * Renew an existing webhook subscription
   *
   * This method validates the user exists before attempting renewal, and automatically
   * deactivates the subscription if the user is not found or inactive.
   *
   * @param subscriptionId - ID of the subscription to renew at Microsoft
   * @param internalUserId - Internal database user ID (MicrosoftUser.id primary key)
   * @returns The renewed subscription data from Microsoft Graph API
   * @throws Error if user not found (after deactivating subscription) or renewal fails
   */
  async renewWebhookSubscription(
    subscriptionId: string,
    internalUserId: number,
  ): Promise<Subscription> {
    const correlationId = `renew-${subscriptionId}-${Date.now()}`;

    try {
      this.logger.log(
        `[${correlationId}] Attempting to renew subscription ${subscriptionId} for user ${internalUserId}`
      );

      // GUARD: Validate user exists and is active
      const user = await this.microsoftUserRepository.findOne({
        where: { id: internalUserId, isActive: true }
      });

      if (!user) {
        // User doesn't exist or inactive - deactivate subscription to prevent future errors
        this.logger.warn(
          `[${correlationId}] User ${internalUserId} not found or inactive. ` +
          `Deactivating subscription ${subscriptionId}`
        );

        await this.webhookSubscriptionRepository.deactivateSubscription(
          subscriptionId
        );

        throw new Error(
          `Cannot renew subscription ${subscriptionId}: ` +
          `User ${internalUserId} not found or inactive. ` +
          `Subscription has been automatically deactivated.`
        );
      }

      this.logger.debug(
        `[${correlationId}] User ${internalUserId} validated successfully`
      );

      // Get access token (handles refresh automatically via getUserAccessToken)
      const accessToken = await this.microsoftAuthService.getUserAccessToken({
        internalUserId
      });

      this.logger.debug(`[${correlationId}] Access token obtained`);

      // Set new expiration date (max 3 days from now per Microsoft limits)
      const expirationDateTime = new Date();
      expirationDateTime.setHours(expirationDateTime.getHours() + 72);

      const renewalData = {
        expirationDateTime: expirationDateTime.toISOString(),
      };

      this.logger.debug(
        `[${correlationId}] Calling Microsoft Graph API to renew subscription`
      );

      // Make the request to Microsoft Graph API to renew the subscription with retry
      const response = await executeGraphApiCall(
        () => axios.patch<Subscription>(
          `${this.graphApiBaseUrl}/subscriptions/${subscriptionId}`,
          renewalData,
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
              "Content-Type": "application/json",
              "Prefer": 'IdType="ImmutableId"',
            },
          }
        ),
        {
          logger: this.logger,
          resourceName: `renew webhook subscription ${subscriptionId} for user ${internalUserId}`,
          maxRetries: 7,
        }
      );

      if (!response?.data) {
        throw new Error('Subscription renewal returned null response');
      }

      this.logger.debug(
        `[${correlationId}] Microsoft Graph API returned status: ${response.status}`
      );

      // Update the expiration date in our local database
      if (response.data.expirationDateTime) {
        await this.webhookSubscriptionRepository.updateSubscriptionExpiration(
          subscriptionId,
          new Date(response.data.expirationDateTime)
        );

        this.logger.debug(
          `[${correlationId}] Updated local database with new expiration: ${response.data.expirationDateTime}`
        );
      }

      this.logger.log(
        `[${correlationId}] Successfully renewed subscription ${subscriptionId}. ` +
        `New expiration: ${response.data.expirationDateTime}`
      );

      return response.data;

    } catch (error: unknown) {
      const errorMessage = error instanceof Error ? error.message : "Unknown error";

      // Special handling for Microsoft API errors
      if (axios.isAxiosError(error)) {
        const statusCode = error.response?.status;

        // Subscription no longer exists at Microsoft — deactivate and attempt re-creation
        if (statusCode === 404) {
          const outcome = await this.handleSubscriptionNotFoundAtMicrosoft(
            subscriptionId,
            internalUserId,
            'renewal_404',
            correlationId,
          );
          if (outcome === 'recreated') {
            // The old subscription has been replaced by a fresh one.
            return {} as Subscription;
          }
          throw new Error(
            `Subscription ${subscriptionId} not found at Microsoft. ` +
            `Subscription has been deactivated and re-creation failed.`
          );
        }

        // User token issues (401, 403) — not transient, deactivate to stop zombie renewals
        if (statusCode === 401 || statusCode === 403) {
          this.logger.error(
            `[${correlationId}] Authentication failed for subscription ${subscriptionId}. ` +
            `Status: ${statusCode}, Response: ${JSON.stringify(error.response?.data)}. ` +
            `Deactivating subscription to prevent repeated failures.`
          );

          await this.webhookSubscriptionRepository.deactivateSubscription(
            subscriptionId
          );

          this.eventEmitter.emit(OutlookEventTypes.SUBSCRIPTION_AUTH_FAILED, {
            subscriptionId,
            userId: internalUserId,
            statusCode,
          });
        }

        // Rate limiting (429)
        if (statusCode === 429) {
          this.logger.warn(
            `[${correlationId}] Rate limited by Microsoft Graph API for subscription ${subscriptionId}`
          );
        }
      }

      this.logger.error(
        `[${correlationId}] Failed to renew subscription ${subscriptionId}: ${errorMessage}`
      );

      throw new Error(`Failed to renew webhook subscription: ${errorMessage}`);
    }
  }

  /**
   * Delete a calendar webhook subscription
   *
   * Deletes the subscription at Microsoft Graph API and deactivates it locally.
   * Supports both external user IDs (from host app) and internal database IDs.
   *
   * @param subscriptionId - ID of the subscription to delete at Microsoft
   * @param userId - User ID (can be external string or internal number)
   * @returns True if deletion was successful
   * @throws Error if user not found or deletion fails (except 404)
   */
  async deleteWebhookSubscription(
    subscriptionId: string,
    userId: string | number,
  ): Promise<boolean> {
    const correlationId = `delete-sub-${subscriptionId}-${Date.now()}`;

    try {
      this.logger.log(
        `[${correlationId}] Deleting calendar subscription ${subscriptionId} for user ${userId}`
      );

      const internalUserId = await this.userIdConverter.toInternalUserId(userId);

      // Get access token (including inactive users since we need to clean up their subscriptions)
      const accessToken = await this.microsoftAuthService.getUserAccessToken({
        internalUserId,
        includeInactive: true
      });

      this.logger.debug(`[${correlationId}] Access token obtained`);

      // Delete subscription at Microsoft
      this.logger.debug(
        `[${correlationId}] Calling Microsoft Graph API to delete subscription`
      );

      await this.deleteSubscription(subscriptionId, accessToken);

      this.logger.log(
        `[${correlationId}] Successfully deleted subscription ${subscriptionId} at Microsoft`
      );

      // Remove the subscription from our database (soft delete)
      await this.webhookSubscriptionRepository.deactivateSubscription(
        subscriptionId
      );

      this.logger.debug(
        `[${correlationId}] Deactivated subscription in local database`
      );

      // Check if user has OTHER active subscriptions before marking inactive
      const otherActiveSubscriptions = await this.webhookSubscriptionRepository.count({
        where: {
          userId: internalUserId,
          isActive: true,
          subscriptionId: Not(subscriptionId)
        }
      });

      // Only mark user inactive if this was the LAST subscription
      if (otherActiveSubscriptions === 0) {
        const updateCriteria = typeof userId === 'string' ? { externalUserId: userId } : { id: userId };
        await this.microsoftUserRepository.update(
          updateCriteria,
          { isActive: false }
        );

        this.logger.debug(
          `[${correlationId}] Marked Microsoft user as inactive - no other subscriptions remain`
        );
      } else {
        this.logger.debug(
          `[${correlationId}] User still has ${otherActiveSubscriptions} other active subscription(s), keeping user active`
        );
      }

      this.logger.log(
        `[${correlationId}] Successfully deleted calendar subscription ${subscriptionId}`
      );

      return true;

    } catch (error: unknown) {
      const errorMessage = error instanceof Error ? error.message : "Unknown error";

      // If we get a 404, the subscription doesn't exist anymore at Microsoft,
      // so we should still remove it from our database
      if (axios.isAxiosError(error) && error.response?.status === 404) {
        this.logger.log(
          `[${correlationId}] Subscription ${subscriptionId} not found at Microsoft, ` +
          `cleaning up local database`
        );

        await this.webhookSubscriptionRepository.deactivateSubscription(
          subscriptionId
        );

        this.logger.log(
          `[${correlationId}] Successfully cleaned up orphaned subscription ${subscriptionId}`
        );

        return true;
      }

      this.logger.error(
        `[${correlationId}] Failed to delete subscription ${subscriptionId}: ${errorMessage}`
      );

      throw new Error(`Failed to delete webhook subscription: ${errorMessage}`);
    }
  }

  // ─── Local database queries ─────────────────────────────────────────

  /**
   * Get a subscription from the local database by subscription ID
   */
  async getSubscription(subscriptionId: string): Promise<OutlookWebhookSubscription | null> {
    return this.webhookSubscriptionRepository.findBySubscriptionId(subscriptionId);
  }

  /**
   * Get active webhook subscription for a user
   * @param externalUserId - External user ID from host application
   * @returns Subscription ID if active subscription exists, null otherwise
   */
  async getActiveSubscriptionForUser(externalUserId: string): Promise<string | null> {
    try {
      // Convert external to internal ID
      const internalUserId = await this.userIdConverter.externalToInternal(externalUserId, {cache: false});

      // A CORRUPTED user needs to re-authenticate — report as not connected so
      // the UI prompts a fresh OAuth flow instead of offering "Disconnect".
      const user = await this.microsoftUserRepository.findOne({
        select: ['id', 'status'],
        where: { id: internalUserId },
      });
      if (!user || user.status === MicrosoftUserStatus.CORRUPTED) {
        this.logger.debug(
          `[getActiveSubscriptionForUser] User ${externalUserId} is missing or CORRUPTED — reporting as not connected`
        );
        return null;
      }

      this.logger.log(`[getActiveSubscriptionForUser] Getting active subscription for user ${externalUserId} (internalUserId: ${internalUserId})`);
      const subscription = await this.webhookSubscriptionRepository.findActiveByUserId(internalUserId);
      this.logger.log(`[getActiveSubscriptionForUser] Found subscription: ${subscription?.subscriptionId}`);
      return subscription?.subscriptionId ?? null;
    } catch {
      // User may not have connected Microsoft account yet - this is not an error
      this.logger.debug(`No active subscription for user ${externalUserId}`);
      return null;
    }
  }

  /**
   * Update the last notification timestamp for a subscription.
   * Fire-and-forget safe — errors are logged internally and never thrown.
   * @param subscriptionId - The subscription ID to update
   */
  trackNotificationReceived(subscriptionId: string): void {
    this.webhookSubscriptionRepository
      .updateLastNotificationAt(subscriptionId)
      .catch((err: unknown) => {
        this.logger.warn(
          `Failed to update lastNotificationAt for subscription ${subscriptionId}: ${err instanceof Error ? err.message : 'Unknown error'}`
        );
      });
  }

  // ─── Cleanup operations ─────────────────────────────────────────────

  /**
   * Cleanup subscriptions with optional filtering
   * Continues on individual deletion failures to ensure maximum cleanup
   * @param options - Cleanup options including access token and optional filter
   * @returns Result summary with counts and errors
   */
  async cleanupSubscriptions(
    options: SubscriptionCleanupOptions,
  ): Promise<SubscriptionCleanupResult> {
    const { accessToken, filter } = options;

    const result: SubscriptionCleanupResult = {
      totalFound: 0,
      successfullyDeleted: 0,
      failedToDelete: 0,
      deletedSubscriptionIds: [],
      errors: [],
    };

    try {
      this.logger.log('🧹 Starting subscription cleanup');

      let subscriptions = await this.getActiveSubscriptions(accessToken);

      // Apply filter if provided
      if (filter) {
        subscriptions = subscriptions.filter(filter);
      }

      result.totalFound = subscriptions.length;

      if (subscriptions.length === 0) {
        this.logger.log('No subscriptions found to clean up');
        return result;
      }

      this.logger.log(`Found ${subscriptions.length} subscription(s) to delete`);

      // Delete each subscription individually, continue on errors
      for (const subscription of subscriptions) {
        try {
          await this.deleteSubscription(subscription.id, accessToken);
          result.successfullyDeleted++;
          result.deletedSubscriptionIds.push(subscription.id);
        } catch (deleteError: unknown) {
          const message = deleteError instanceof Error ? deleteError.message : 'Unknown error';
          result.failedToDelete++;
          result.errors.push({
            subscriptionId: subscription.id,
            error: deleteError instanceof Error ? deleteError.message : 'Unknown error',
          });
          this.logger.warn(`⚠️ Failed to delete subscription ${subscription.id}: ${message}`);
        }
      }

      this.logger.log(
        `🗑️ Cleanup completed: ${result.successfullyDeleted} deleted, ${result.failedToDelete} failed`,
      );
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`❌ Cleanup operation failed: ${message}`);
      throw error;
    }

    return result;
  }

  /**
   * Cleanup subscriptions for a specific client ID
   * @param clientId - The client ID to cleanup subscriptions for
   * @param accessToken - Valid Microsoft access token
   * @returns Result summary
   */
  async cleanupSubscriptionsForClient(
    clientId: number,
    accessToken: string,
  ): Promise<SubscriptionCleanupResult> {
    this.logger.log(`🧹 Cleaning up subscriptions for client ${clientId}`);
    return this.cleanupSubscriptions({
      accessToken,
      filter: (sub) => sub.clientState?.includes(`client_${clientId}_`) || false,
    });
  }

  /**
   * Cleanup subscriptions for a specific user ID
   * @param userId - The user ID to cleanup subscriptions for
   * @param accessToken - Valid Microsoft access token
   * @returns Result summary
   */
  async cleanupSubscriptionsForUser(
    userId: number,
    accessToken: string,
  ): Promise<SubscriptionCleanupResult> {
    this.logger.log(`🧹 Cleaning up subscriptions for user ${userId}`);
    return this.cleanupSubscriptions({
      accessToken,
      filter: (sub) => sub.clientState?.includes(`user_${userId}_`) || false,
    });
  }

  /**
   * Cleanup all active subscriptions for a user + resource, at Microsoft Graph
   * and in the local DB. Leaves subs for other resources untouched.
   *
   * @param externalUserId - External user ID from the host application
   * @param resource - Graph resource string (e.g. '/me/events', '/me/messages')
   * @returns Result summary from the Graph-side cleanup pass
   */
  async cleanupSubscriptionsForUserAndResource(
    externalUserId: string,
    resource: string,
  ): Promise<SubscriptionCleanupResult> {
    const internalUserId = await this.userIdConverter.externalToInternal(
      externalUserId,
      { cache: false },
    );
    const accessToken = await this.microsoftAuthService.getUserAccessToken({
      internalUserId,
      cache: false,
    });

    this.logger.log(
      `🧹 Cleaning up ${resource} subscriptions for user ${internalUserId}`,
    );

    const result = await this.cleanupSubscriptions({
      accessToken,
      filter: (sub) =>
        sub.resource === resource &&
        (sub.clientState?.includes(`user_${internalUserId}_`) || false),
    });

    const dbRows = await this.webhookSubscriptionRepository
      .findAllActiveByUserIdAndResource(internalUserId, resource);
    for (const row of dbRows) {
      await this.webhookSubscriptionRepository.deactivateSubscription(
        row.subscriptionId,
      );
    }

    return result;
  }

  /**
   * Revoke Microsoft OAuth tokens
   * @param refreshToken - The refresh token to revoke
   */
  async revokeTokens(refreshToken: string): Promise<void> {
    try {
      if (!refreshToken) {
        this.logger.warn('⚠️ No refresh token available for revocation');
        return;
      }

      await axios.post(
        `${this.msAuthUrl}/logout`,
        new URLSearchParams({
          token: refreshToken,
          token_type_hint: 'refresh_token',
        }),
        {
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
          timeout: 10000,
        },
      );

      this.logger.log('✅ Microsoft tokens revoked successfully');
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      this.logger.warn(`⚠️ Failed to revoke Microsoft tokens: ${message}`);
    }
  }

  /**
   * Full cleanup: revoke tokens and delete subscriptions
   * @param refreshToken - The refresh token to revoke
   * @param accessToken - Valid Microsoft access token
   * @param filter - Optional filter for subscriptions to delete
   * @returns Result summary
   */
  async fullCleanup(
    refreshToken: string,
    accessToken: string,
    filter?: SubscriptionFilter,
  ): Promise<SubscriptionCleanupResult> {
    this.logger.log('🔄 Starting full cleanup (tokens + subscriptions)');

    // 1. Cleanup subscriptions first
    const result = await this.cleanupSubscriptions({
      accessToken,
      filter,
    });

    // 2. Revoke tokens
    await this.revokeTokens(refreshToken);

    this.logger.log('✅ Full cleanup completed');
    return result;
  }

  // ─── Scheduled jobs ─────────────────────────────────────────────────

  /**
   * Scheduled job that verifies active subscriptions still exist at Microsoft
   * and detects stale subscriptions that stopped receiving notifications.
   *
   * Runs every 6 hours. Work is split into two passes:
   * 1. Stale-notification pass (in-memory, no Graph calls) — emits LIFECYCLE_MISSED
   *    for any sub whose lastNotificationAt is older than 24 h.
   * 2. Graph-verify pass (only subs expiring within 24 h) — grouped by user, batched
   *    via /$batch (20 per call), rate-limited per user. Handles 404 re-creation and
   *    "<12 h to expiry" forced renewal.
   */
  @Cron('0 */6 * * *')
  async verifySubscriptionHealth(): Promise<void> {
    const correlationId = `health-check-${Date.now()}`;

    try {
      const allActiveSubs = await this.webhookSubscriptionRepository.findActiveSubscriptions();

      if (allActiveSubs.length === 0) {
        this.logger.debug(`[${correlationId}] No active subscriptions to verify`);
        return;
      }

      // Skip subs whose owning user is CORRUPTED — any outbound Graph work for
      // them would just trigger the same invalid_grant and we've already signalled
      // the host app to prompt re-auth.
      const corruptedUserIds = await this.findCorruptedUserIds(allActiveSubs);
      const activeSubs = corruptedUserIds.size > 0
        ? allActiveSubs.filter((sub) => !corruptedUserIds.has(sub.userId))
        : allActiveSubs;

      const skipped = allActiveSubs.length - activeSubs.length;
      if (skipped > 0) {
        this.logger.log(
          `[${correlationId}] Skipping ${skipped} sub(s) owned by ${corruptedUserIds.size} CORRUPTED user(s)`
        );
      }

      const expiringSoon = this.filterExpiringSoon(activeSubs, HEALTH_CHECK_WINDOW_HOURS);

      this.logger.log(
        `[${correlationId}] Health check: ${activeSubs.length} active, ${expiringSoon.length} expiring within ${HEALTH_CHECK_WINDOW_HOURS}h`
      );

      const counters: HealthCheckCounters = { verified: 0, recreated: 0, renewed: 0, staleDetected: 0, failed: 0 };

      counters.staleDetected = this.detectStaleSubscriptions(activeSubs, correlationId);

      if (expiringSoon.length > 0) {
        const byUser = this.groupByUserId(expiringSoon);
        for (const [userId, userSubs] of byUser) {
          const perUser = await this.verifySubscriptionsForUser(userId, userSubs, correlationId);
          counters.verified += perUser.verified;
          counters.recreated += perUser.recreated;
          counters.renewed += perUser.renewed;
          counters.failed += perUser.failed;
        }
      }

      this.logger.log(
        `[${correlationId}] Health check complete: ${counters.verified} verified, ${counters.recreated} recreated, ${counters.renewed} renewed, ${counters.staleDetected} stale-detected, ${counters.failed} failed`
      );
    } catch (error) {
      this.logger.error(`[${correlationId}] Subscription health check job failed:`, error);
    }
  }

  /**
   * Return the set of user IDs whose MicrosoftUser.status === CORRUPTED among
   * the owners of the given subs. One indexed query, no N+1.
   */
  private async findCorruptedUserIds(
    subs: OutlookWebhookSubscription[],
  ): Promise<Set<number>> {
    const userIds = Array.from(new Set(subs.map((s) => s.userId)));
    if (userIds.length === 0) return new Set();

    const rows = await this.microsoftUserRepository.find({
      select: ['id'],
      where: { id: In(userIds), status: MicrosoftUserStatus.CORRUPTED },
    });
    return new Set(rows.map((u) => u.id));
  }

  // ─── Health check helpers ───────────────────────────────────────────

  /**
   * Pure filter: returns subs whose expirationDateTime is within `hoursAhead` of now.
   */
  private filterExpiringSoon(
    subs: OutlookWebhookSubscription[],
    hoursAhead: number,
  ): OutlookWebhookSubscription[] {
    const threshold = new Date(Date.now() + hoursAhead * 3600 * 1000);
    return subs.filter((sub) => sub.expirationDateTime < threshold);
  }

  /**
   * Pure grouping: bucket subscriptions by their internal userId.
   */
  private groupByUserId(
    subs: OutlookWebhookSubscription[],
  ): Map<number, OutlookWebhookSubscription[]> {
    const byUser = new Map<number, OutlookWebhookSubscription[]>();
    for (const sub of subs) {
      const bucket = byUser.get(sub.userId);
      if (bucket) bucket.push(sub);
      else byUser.set(sub.userId, [sub]);
    }
    return byUser;
  }

  /**
   * Emit LIFECYCLE_MISSED for every sub whose last notification is older than 24 h.
   * Returns the count of stale subs detected. No Graph calls.
   */
  private detectStaleSubscriptions(
    subs: OutlookWebhookSubscription[],
    correlationId: string,
  ): number {
    const staleThreshold = new Date(Date.now() - STALE_NOTIFICATION_THRESHOLD_HOURS * 3600 * 1000);
    let staleDetected = 0;

    for (const sub of subs) {
      if (!sub.lastNotificationAt || sub.lastNotificationAt >= staleThreshold) continue;

      this.logger.warn(
        `[${correlationId}] Subscription ${sub.subscriptionId} has not received notifications since ${sub.lastNotificationAt.toISOString()}. Emitting LIFECYCLE_MISSED for delta sync.`
      );

      this.eventEmitter.emit(OutlookEventTypes.LIFECYCLE_MISSED, {
        subscriptionId: sub.subscriptionId,
        userId: sub.userId,
        reason: 'health_check_stale',
      });

      staleDetected++;
    }

    return staleDetected;
  }

  /**
   * Resolve one token for this user, then verify their subs in /$batch chunks of 20.
   * A failure to obtain the token counts every sub for this user as failed.
   */
  private async verifySubscriptionsForUser(
    userId: number,
    userSubs: OutlookWebhookSubscription[],
    correlationId: string,
  ): Promise<HealthCheckCounters> {
    const totals: HealthCheckCounters = { verified: 0, recreated: 0, renewed: 0, staleDetected: 0, failed: 0 };

    let externalUserId: string;
    let accessToken: string;
    try {
      externalUserId = await this.userIdConverter.internalToExternal(userId);
      accessToken = await this.microsoftAuthService.getUserAccessToken({ internalUserId: userId });
    } catch (error) {
      totals.failed += userSubs.length;
      this.logger.error(
        `[${correlationId}] Token fetch failed for user ${userId}; skipping ${userSubs.length} sub(s):`,
        error
      );
      return totals;
    }

    for (let offset = 0; offset < userSubs.length; offset += GRAPH_BATCH_SIZE) {
      const chunk = userSubs.slice(offset, offset + GRAPH_BATCH_SIZE);
      const chunkTotals = await this.verifySubscriptionsBatch(chunk, userId, externalUserId, accessToken, correlationId);
      totals.verified += chunkTotals.verified;
      totals.recreated += chunkTotals.recreated;
      totals.renewed += chunkTotals.renewed;
      totals.failed += chunkTotals.failed;
    }

    return totals;
  }

  /**
   * Issue one /$batch POST for up to 20 subs belonging to the same user.
   * Dispatches each inner response to `handleBatchResponseItem`.
   */
  private async verifySubscriptionsBatch(
    chunk: OutlookWebhookSubscription[],
    userId: number,
    externalUserId: string,
    accessToken: string,
    correlationId: string,
  ): Promise<HealthCheckCounters> {
    const totals: HealthCheckCounters = { verified: 0, recreated: 0, renewed: 0, staleDetected: 0, failed: 0 };

    await this.rateLimiter.acquirePermit(externalUserId);

    const batchPayload: BatchRequestPayload = {
      requests: chunk.map((sub, index) => ({
        id: String(index),
        method: 'GET',
        url: `/subscriptions/${sub.subscriptionId}`,
        headers: { 'Prefer': 'IdType="ImmutableId"' },
      })),
    };

    try {
      const response = await executeGraphApiCall(
        () => axios.post<BatchResponsePayload<Subscription>>(
          `${this.graphApiBaseUrl}/$batch`,
          batchPayload,
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
              'Content-Type': 'application/json',
            },
          }
        ),
        {
          logger: this.logger,
          resourceName: `verify batch (user ${userId}, ${chunk.length} subs)`,
          maxRetries: 2,
        }
      );

      if (!response?.data) {
        totals.failed += chunk.length;
        this.logger.error(`[${correlationId}] Batch verify returned null for user ${userId}`);
        return totals;
      }

      for (const item of response.data.responses) {
        const sub = chunk[parseInt(item.id, 10)] as OutlookWebhookSubscription | undefined;
        if (!sub) continue;
        const itemTotals = await this.handleBatchResponseItem(sub, item, correlationId);
        totals.verified += itemTotals.verified;
        totals.recreated += itemTotals.recreated;
        totals.renewed += itemTotals.renewed;
        totals.failed += itemTotals.failed;
      }
    } catch (error) {
      totals.failed += chunk.length;
      this.logger.error(
        `[${correlationId}] Batch verify failed for user ${userId} (${chunk.length} subs):`,
        error
      );
    }

    return totals;
  }

  /**
   * Branch on one inner /$batch response: 200 (possibly renew), 404 (re-create), other (log).
   */
  private async handleBatchResponseItem(
    sub: OutlookWebhookSubscription,
    item: BatchResponse<Subscription>,
    correlationId: string,
  ): Promise<HealthCheckCounters> {
    const totals: HealthCheckCounters = { verified: 0, recreated: 0, renewed: 0, staleDetected: 0, failed: 0 };

    if (item.status === 200) {
      totals.verified = 1;
      if (await this.maybeRenewIfExpiringSoon(sub, item.body, correlationId)) {
        totals.renewed = 1;
      }
      return totals;
    }

    if (item.status === 404) {
      const result = await this.handleSubscriptionNotFoundAtMicrosoft(sub.subscriptionId, sub.userId, 'health_check_404', correlationId);
      if (result === 'recreated') totals.recreated = 1;
      else totals.failed = 1;
      return totals;
    }

    totals.failed = 1;
    this.logger.error(
      `[${correlationId}] Unexpected status ${item.status} for sub ${sub.subscriptionId}: ${JSON.stringify(item.body)}`
    );
    return totals;
  }

  /**
   * If the Graph-side expirationDateTime is within 12 h, trigger a renewal.
   * Returns true iff a renewal was attempted.
   */
  private async maybeRenewIfExpiringSoon(
    sub: OutlookWebhookSubscription,
    graphSub: Subscription,
    correlationId: string,
  ): Promise<boolean> {
    if (!graphSub.expirationDateTime) return false;

    const expiration = new Date(graphSub.expirationDateTime);
    const renewalThreshold = new Date(Date.now() + FORCE_RENEWAL_WINDOW_HOURS * 3600 * 1000);
    if (expiration >= renewalThreshold) return false;

    this.logger.log(
      `[${correlationId}] Subscription ${sub.subscriptionId} expires within ${FORCE_RENEWAL_WINDOW_HOURS}h. Forcing renewal.`
    );
    try {
      await this.renewWebhookSubscription(sub.subscriptionId, sub.userId);
    } catch (error) {
      // Swallow — renewWebhookSubscription already logs + emits events on failure.
      this.logger.warn(
        `[${correlationId}] Forced renewal of ${sub.subscriptionId} failed: ${error instanceof Error ? error.message : 'Unknown error'}`
      );
    }
    return true;
  }

  /**
   * Deactivate locally and attempt re-create. Emits SUBSCRIPTION_RECREATED or
   * SUBSCRIPTION_RECREATION_FAILED. Used by both the renewal 404 handler and the
   * health-check batch response handler.
   */
  private async handleSubscriptionNotFoundAtMicrosoft(
    subscriptionId: string,
    internalUserId: number,
    reason: 'renewal_404' | 'health_check_404',
    correlationId: string,
  ): Promise<'recreated' | 'failed'> {
    this.logger.warn(
      `[${correlationId}] Subscription ${subscriptionId} not found at Microsoft. Deactivating and attempting re-creation.`
    );

    await this.webhookSubscriptionRepository.deactivateSubscription(subscriptionId);

    try {
      const externalUserId = await this.userIdConverter.internalToExternal(internalUserId);
      await this.createWebhookSubscription(externalUserId);

      this.logger.log(
        `[${correlationId}] Successfully re-created subscription for user ${internalUserId} after 404`
      );
      this.eventEmitter.emit(OutlookEventTypes.SUBSCRIPTION_RECREATED, {
        subscriptionId,
        userId: internalUserId,
        reason,
      });
      return 'recreated';
    } catch (recreateError) {
      const recreateMsg = recreateError instanceof Error ? recreateError.message : 'Unknown error';
      this.logger.error(
        `[${correlationId}] Failed to re-create subscription for user ${internalUserId}: ${recreateMsg}`
      );
      this.eventEmitter.emit(OutlookEventTypes.SUBSCRIPTION_RECREATION_FAILED, {
        subscriptionId,
        userId: internalUserId,
        reason,
        error: recreateMsg,
      });
      return 'failed';
    }
  }
}
