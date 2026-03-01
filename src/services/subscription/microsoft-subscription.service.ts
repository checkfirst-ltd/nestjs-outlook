import { Injectable, Logger } from '@nestjs/common';
import axios from 'axios';
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

/**
 * Centralized service for managing Microsoft Graph API subscriptions
 * Handles creation, deletion, and cleanup of webhook subscriptions
 * for calendar, email, and other Microsoft integrations
 */
@Injectable()
export class MicrosoftSubscriptionService {
  private readonly logger = new Logger(MicrosoftSubscriptionService.name);
  private readonly graphApiBaseUrl = 'https://graph.microsoft.com/v1.0';
  private readonly msAuthUrl = 'https://login.microsoftonline.com/common/oauth2/v2.0';

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
          },
          timeout: 10000,
        }),
        {
          logger: this.logger,
          resourceName: 'get active subscriptions',
          maxRetries: 3,
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
          },
          timeout: 10000,
        }),
        {
          logger: this.logger,
          resourceName: `delete subscription ${subscriptionId}`,
          maxRetries: 3,
          return404AsNull: true,
        }
      );
      this.logger.log(`✅ Deleted subscription ${subscriptionId} at Microsoft`);
    } catch (error: unknown) {
      throw error;
    }
  }

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
}
