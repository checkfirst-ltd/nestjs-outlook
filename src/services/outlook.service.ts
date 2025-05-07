import { Injectable, Logger, Inject, forwardRef } from '@nestjs/common';
import { EventEmitter2 } from '@nestjs/event-emitter';
import { Client } from '@microsoft/microsoft-graph-client';
import axios from 'axios';
import {
  Event,
  Calendar,
  Subscription,
  ChangeNotification,
} from '../types';
import { MicrosoftAuthService } from './microsoft-auth.service';
import { TokenResponse } from '../interfaces/outlook/token-response.interface';
import { Cron, CronExpression } from '@nestjs/schedule';
import { OutlookWebhookSubscriptionRepository } from '../repositories/outlook-webhook-subscription.repository';
import { OutlookResourceData } from '../dto/outlook-webhook-notification.dto';
import { MICROSOFT_CONFIG } from '../constants';
import { MicrosoftOutlookConfig } from '../interfaces/config/outlook-config.interface';
import { OutlookEventTypes } from '../event-types.enum';

@Injectable()
export class OutlookService {
  private readonly logger = new Logger(OutlookService.name);

  constructor(
    @Inject(forwardRef(() => MicrosoftAuthService))
    private readonly microsoftAuthService: MicrosoftAuthService,
    private readonly webhookSubscriptionRepository: OutlookWebhookSubscriptionRepository,
    private readonly eventEmitter: EventEmitter2,
    @Inject(MICROSOFT_CONFIG)
    private readonly microsoftConfig: MicrosoftOutlookConfig,
  ) {}

  /**
   * Get the user's default calendar ID
   * @param accessToken - Access token for Microsoft Graph API
   * @returns The default calendar ID
   */
  async getDefaultCalendarId(accessToken: string): Promise<string> {
    try {
      // Using axios for direct API call
      const response = await axios.get<Calendar>('https://graph.microsoft.com/v1.0/me/calendar', {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.data.id) {
        throw new Error('Failed to retrieve calendar ID');
      }

      return response.data.id;
    } catch (error) {
      this.logger.error('Error getting default calendar ID:', error);
      throw new Error('Failed to get calendar ID from Microsoft');
    }
  }

  /**
   * Creates an event in the user's Outlook calendar
   * @param event - Microsoft Graph Event object with event details
   * @param accessToken - Access token for Microsoft Graph API
   * @param refreshToken - Refresh token for Microsoft Graph API
   * @param tokenExpiry - Expiry date of the access token
   * @param userId - User ID associated with the calendar
   * @param calendarId - Calendar ID where the event will be created
   * @returns The created event data and refreshed token data if tokens were refreshed
   */
  async createEvent(
    event: Partial<Event>,
    accessToken: string,
    refreshToken: string,
    tokenExpiry: string | undefined,
    userId: number,
    calendarId: string,
  ): Promise<{ event: Event; tokensRefreshed: boolean; refreshedTokens?: TokenResponse }> {
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
              userId,
              calendarId,
            );

            // Update the access token
            currentAccessToken = refreshedTokens.access_token;
            currentRefreshToken = refreshedTokens.refresh_token;
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

      // Create the event
      const createdEvent = await client
        .api(`/me/calendars/${calendarId}/events`)
        .post(event) as Event;

      // Return both the event and token refresh information
      return {
        event: createdEvent,
        tokensRefreshed,
        refreshedTokens,
      };
    } catch (error: unknown) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to create Outlook calendar event: ${errorMessage}`);
      throw new Error(`Failed to create Outlook calendar event: ${errorMessage}`);
    }
  }

  /**
   * Create a webhook subscription to receive notifications for calendar events
   * @param userId - User ID
   * @param accessToken - Access token for Microsoft Graph API
   * @param refreshToken - Refresh token for Microsoft Graph API
   * @returns The created subscription data
   */
  async createWebhookSubscription(
    userId: number,
    accessToken: string,
    refreshToken: string,
  ): Promise<Subscription> {
    try {
      // Set expiration date (max 3 days as per Microsoft documentation)
      const expirationDateTime = new Date();
      expirationDateTime.setHours(expirationDateTime.getHours() + 72); // 3 days from now

      const appUrl = this.microsoftConfig.backendBaseUrl || 'http://localhost:3000';
      const basePath = this.microsoftConfig.basePath;
      const basePathUrl = basePath ? `${appUrl}/${basePath}` : appUrl;

      // Create subscription payload with proper URL encoding
      const notificationUrl = `${basePathUrl}/outlook/webhook`;

      // Create subscription payload
      const subscriptionData = {
        changeType: 'created,updated,deleted',
        notificationUrl,
        // Add lifecycleNotificationUrl for increased reliability
        lifecycleNotificationUrl: notificationUrl,
        resource: '/me/events',
        expirationDateTime: expirationDateTime.toISOString(),
        clientState: `user_${String(userId)}_${Math.random().toString(36).substring(2, 15)}`,
      };

      this.logger.debug(`Creating webhook subscription with notificationUrl: ${notificationUrl}`);

      this.logger.debug(`Subscription data: ${JSON.stringify(subscriptionData)}`);
      // Create the subscription with Microsoft Graph API
      const response = await axios.post<Subscription>(
        'https://graph.microsoft.com/v1.0/subscriptions',
        subscriptionData,
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/json',
          },
        },
      );

      this.logger.log(`Created webhook subscription ${String(response.data.id)} for user ${String(userId)}`);

      // Save the subscription to the database
      await this.webhookSubscriptionRepository.saveSubscription({
        subscriptionId: response.data.id,
        userId,
        resource: response.data.resource,
        changeType: response.data.changeType,
        clientState: response.data.clientState || '',
        notificationUrl: response.data.notificationUrl,
        expirationDateTime: response.data.expirationDateTime ? new Date(response.data.expirationDateTime) : new Date(),
        accessToken,
        refreshToken,
      });

      this.logger.debug(
        `Stored subscription with refresh token: ${refreshToken.substring(0, 5)}...`,
      );

      return response.data;
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to create webhook subscription: ${errorMessage}`);
      throw new Error(`Failed to create webhook subscription: ${errorMessage}`);
    }
  }

  /**
   * Renew a webhook subscription before it expires
   * @param subscriptionId - ID of the subscription to renew
   * @param accessToken - Access token for Microsoft Graph API
   * @returns The renewed subscription data
   */
  async renewWebhookSubscription(
    subscriptionId: string,
    accessToken: string,
  ): Promise<Subscription> {
    try {
      // Set a new expiration date (max 3 days as per Microsoft documentation)
      const expirationDateTime = new Date();
      expirationDateTime.setHours(expirationDateTime.getHours() + 72); // 3 days from now

      // Update the subscription with Microsoft Graph API
      const response = await axios.patch<Subscription>(
        `https://graph.microsoft.com/v1.0/subscriptions/${subscriptionId}`,
        {
          expirationDateTime: expirationDateTime.toISOString(),
        },
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/json',
          },
        },
      );

      this.logger.log(
        `Renewed webhook subscription ${subscriptionId} until ${expirationDateTime.toISOString()}`,
      );

      // Update the subscription in the database
      await this.webhookSubscriptionRepository.updateSubscriptionExpiration(
        subscriptionId,
        expirationDateTime,
        accessToken,
      );

      return response.data;
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to renew webhook subscription: ${errorMessage}`);
      throw new Error(`Failed to renew webhook subscription: ${errorMessage}`);
    }
  }

  /**
   * Delete a webhook subscription
   * @param subscriptionId - ID of the subscription to delete
   * @param accessToken - Access token for Microsoft Graph API
   * @returns Success status
   */
  async deleteWebhookSubscription(subscriptionId: string, accessToken: string): Promise<boolean> {
    try {
      await axios.delete(`https://graph.microsoft.com/v1.0/subscriptions/${subscriptionId}`, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
        },
      });

      this.logger.log(`Deleted webhook subscription ${subscriptionId}`);

      // Deactivate subscription in the database
      await this.webhookSubscriptionRepository.deactivateSubscription(subscriptionId);

      return true;
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to delete webhook subscription: ${errorMessage}`);

      // If error is 404 (subscription not found), deactivate it locally
      if (axios.isAxiosError(error) && error.response?.status === 404) {
        await this.webhookSubscriptionRepository.deactivateSubscription(subscriptionId);
        this.logger.log(
          `Subscription ${subscriptionId} not found on Microsoft servers, deactivated locally`,
        );
        return true;
      }

      throw new Error(`Failed to delete webhook subscription: ${errorMessage}`);
    }
  }

  /**
   * Scheduled task to renew expiring subscriptions
   * Runs every day at 2 AM to check and renew subscriptions expiring within 24 hours
   */
  @Cron(CronExpression.EVERY_DAY_AT_2AM)
  async renewSubscriptions(): Promise<void> {
    this.logger.log('Running scheduled task to check for expiring subscriptions');

    try {
      // Find subscriptions that will expire within 24 hours
      const expiringSubscriptions =
        await this.webhookSubscriptionRepository.findSubscriptionsNeedingRenewal(24);

      if (expiringSubscriptions.length === 0) {
        this.logger.log('No subscriptions need renewal at this time');
        return;
      }

      this.logger.log(
        `Found ${String(expiringSubscriptions.length)} subscriptions expiring within 24 hours`,
      );

      // Renew each subscription
      for (const subscription of expiringSubscriptions) {
        try {
          await this.microsoftAuthService.renewWebhookSubscription(
            subscription.subscriptionId,
            subscription.accessToken,
            subscription.refreshToken,
          );

          this.logger.log(
            `Successfully renewed subscription ${subscription.subscriptionId} for user ${String(subscription.userId)}`,
          );
        } catch (error) {
          const errorMessage = error instanceof Error ? error.message : 'Unknown error';
          this.logger.error(
            `Failed to renew subscription ${subscription.subscriptionId}: ${errorMessage}`,
          );
          // Continue with next subscription even if one fails
        }
      }

      this.logger.log('Finished renewing webhook subscriptions');
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Error in scheduled task to renew subscriptions: ${errorMessage}`);
    }
  }

  /**
   * Handle webhook notifications from Microsoft Graph API
   * Supports event deletion, creation, and updates
   *
   * @param notification - The notification item from Microsoft Graph
   * @returns Result of the operation
   */
  async handleOutlookWebhook(
    notificationItem: ChangeNotification,
  ): Promise<{ success: boolean; message: string }> {
    try {
      // Cast the resourceData to our class
      const resourceData = notificationItem.resourceData as OutlookResourceData;

      if (!resourceData.id) {
        throw new Error('No event ID found in the webhook notification payload');
      }

      this.logger.log(
        `Processing Outlook event ${String(notificationItem.changeType)} for event ID: ${resourceData.id}`,
      );

      // Handle different event types
      switch (notificationItem.changeType) {
        case 'deleted':
          // Emit an event that will be caught by the calendar service
          await Promise.resolve(
            this.eventEmitter.emit(OutlookEventTypes.EVENT_DELETED, resourceData),
          );
          return {
            success: true,
            message: `Event deletion notification processed for ID: ${resourceData.id}`,
          };

        case 'created':
          await Promise.resolve(
            this.eventEmitter.emit(OutlookEventTypes.EVENT_CREATED, resourceData),
          );
          return {
            success: true,
            message: `Event creation notification processed for ID: ${resourceData.id}`,
          };

        case 'updated':
          await Promise.resolve(
            this.eventEmitter.emit(OutlookEventTypes.EVENT_UPDATED, resourceData),
          );
          return {
            success: true,
            message: `Event update notification processed for ID: ${resourceData.id}`,
          };

        default: {
          const changeType = 'unknown';
          return {
            success: false,
            message: `Notification type '${changeType}' not supported`,
          };
        }
      }
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to process Outlook webhook notification: ${errorMessage}`);
      throw new Error(`Failed to process Outlook webhook notification: ${errorMessage}`);
    }
  }
}
