import { Injectable, Logger, Inject, forwardRef } from '@nestjs/common';
import { EventEmitter2 } from '@nestjs/event-emitter';
import { Client } from '@microsoft/microsoft-graph-client';
import axios from 'axios';
import {
  Event,
  Calendar,
  Subscription,
  ChangeNotification,
} from '../../types';
import { MicrosoftAuthService } from '../auth/microsoft-auth.service';
import { TokenResponse } from '../../interfaces/outlook/token-response.interface';
import { Cron, CronExpression } from '@nestjs/schedule';
import { OutlookWebhookSubscriptionRepository } from '../../repositories/outlook-webhook-subscription.repository';
import { OutlookResourceData } from '../../dto/outlook-webhook-notification.dto';
import { MICROSOFT_CONFIG } from '../../constants';
import { MicrosoftOutlookConfig } from '../../interfaces/config/outlook-config.interface';
import { OutlookEventTypes } from '../../event-types.enum';

@Injectable()
export class CalendarService {
  private readonly logger = new Logger(CalendarService.name);

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
            currentAccessToken = refreshedTokens.access_token || currentAccessToken;
            currentRefreshToken = refreshedTokens.refresh_token || currentRefreshToken;
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
      const notificationUrl = `${basePathUrl}/calendar/webhook`;

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
    } catch (error: unknown) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to create webhook subscription: ${errorMessage}`);
      throw new Error(`Failed to create webhook subscription: ${errorMessage}`);
    }
  }

  /**
   * Renew an existing webhook subscription
   * @param subscriptionId - ID of the subscription to renew
   * @param accessToken - Access token for Microsoft Graph API
   * @returns The renewed subscription data
   */
  async renewWebhookSubscription(
    subscriptionId: string,
    accessToken: string,
  ): Promise<Subscription> {
    try {
      // Set new expiration date (max 3 days from now)
      const expirationDateTime = new Date();
      expirationDateTime.setHours(expirationDateTime.getHours() + 72);

      // Prepare the renewal payload
      const renewalData = {
        expirationDateTime: expirationDateTime.toISOString(),
      };

      // Make the request to Microsoft Graph API to renew the subscription
      const response = await axios.patch<Subscription>(
        `https://graph.microsoft.com/v1.0/subscriptions/${subscriptionId}`,
        renewalData,
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/json',
          },
        },
      );

      // Update the expiration date in our database
      if (response.data.expirationDateTime) {
        await this.webhookSubscriptionRepository.updateSubscriptionExpiration(
          subscriptionId,
          new Date(response.data.expirationDateTime),
        );
      }

      this.logger.log(`Renewed webhook subscription: ${subscriptionId}`);

      return response.data;
    } catch (error: unknown) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to renew webhook subscription: ${errorMessage}`);
      throw new Error(`Failed to renew webhook subscription: ${errorMessage}`);
    }
  }

  /**
   * Delete a webhook subscription
   * @param subscriptionId - ID of the subscription to delete
   * @param accessToken - Access token for Microsoft Graph API
   * @returns True if deletion was successful
   */
  async deleteWebhookSubscription(subscriptionId: string, accessToken: string): Promise<boolean> {
    try {
      // Make the request to Microsoft Graph API to delete the subscription
      await axios.delete(
        `https://graph.microsoft.com/v1.0/subscriptions/${subscriptionId}`,
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/json',
          },
        },
      );

      // Remove the subscription from our database
      await this.webhookSubscriptionRepository.deactivateSubscription(subscriptionId);

      this.logger.log(`Deleted webhook subscription: ${subscriptionId}`);

      return true;
    } catch (error: unknown) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to delete webhook subscription: ${errorMessage}`);

      // If we get a 404, the subscription doesn't exist anymore at Microsoft,
      // so we should remove it from our database
      if (axios.isAxiosError(error) && error.response?.status === 404) {
        await this.webhookSubscriptionRepository.deactivateSubscription(subscriptionId);
        this.logger.log(`Subscription not found, removed from database: ${subscriptionId}`);
        return true;
      }

      throw new Error(`Failed to delete webhook subscription: ${errorMessage}`);
    }
  }

  /**
   * Scheduled job that checks for webhook subscriptions that will expire soon
   * and renews them
   */
  @Cron(CronExpression.EVERY_HOUR)
  async renewSubscriptions(): Promise<void> {
    try {
      // Get subscriptions that expire within the next 24 hours
      const expiringLimit = new Date();
      expiringLimit.setHours(expiringLimit.getHours() + 24);

      const subscriptions = await this.webhookSubscriptionRepository.findSubscriptionsNeedingRenewal(
        24 // hours until expiration
      );

      if (subscriptions.length === 0) {
        this.logger.debug('No subscriptions need renewal');
        return;
      }

      this.logger.log(`Found ${String(subscriptions.length)} subscriptions that need renewal`);

      // Renew each subscription
      for (const subscription of subscriptions) {
        try {
          // Check if we need to refresh the access token first
          let accessToken = subscription.accessToken;

          if (subscription.refreshToken) {
            try {
              const tokenResponse = await this.microsoftAuthService.refreshAccessToken(
                subscription.refreshToken,
                subscription.userId,
              );
              accessToken = tokenResponse.access_token;

              // Create expiration date (3 days from now)
              const newExpirationDate = new Date();
              newExpirationDate.setHours(newExpirationDate.getHours() + 72); // 3 days
              
              // Update the subscription in the repository with the new tokens
              await this.webhookSubscriptionRepository.updateSubscriptionExpiration(
                subscription.subscriptionId,
                newExpirationDate,
                accessToken
              );
            } catch (refreshError) {
              this.logger.error(
                `Failed to refresh token for subscription ${subscription.subscriptionId}:`,
                refreshError,
              );
              continue; // Skip this subscription and try the next one
            }
          }

          // Now renew the subscription with the (possibly refreshed) access token
          await this.renewWebhookSubscription(subscription.subscriptionId, accessToken);
        } catch (error) {
          this.logger.error(
            `Failed to renew subscription ${subscription.subscriptionId}:`,
            error,
          );
          // Continue with the next subscription even if this one failed
        }
      }
    } catch (error) {
      this.logger.error('Error in subscription renewal job:', error);
    }
  }

  /**
   * Handle a webhook notification from Microsoft
   * @param notificationItem - The notification data from Microsoft
   * @returns Success status and message
   */
  async handleOutlookWebhook(
    notificationItem: ChangeNotification,
  ): Promise<{ success: boolean; message: string }> {
    try {
      // Extract necessary information from the notification
      const { subscriptionId, clientState, resource, changeType } = notificationItem;

      this.logger.debug(`Received webhook notification for subscription: ${subscriptionId || 'unknown'}`);
      this.logger.debug(`Resource: ${resource || 'unknown'}, ChangeType: ${String(changeType || 'unknown')}`);

      // Find the subscription in our database to verify it's legitimate
      const subscription = await this.webhookSubscriptionRepository.findBySubscriptionId(
        subscriptionId || '',
      );

      if (!subscription) {
        this.logger.warn(`Unknown subscription ID: ${subscriptionId || 'unknown'}`);
        return { success: false, message: 'Unknown subscription' };
      }

      // Verify the client state for additional security
      if (subscription.clientState && clientState !== subscription.clientState) {
        this.logger.warn('Client state mismatch');
        return { success: false, message: 'Client state mismatch' };
      }

      // Extract the user ID from the client state (should be in format "user_123_randomstring")
      const userId = subscription.userId;

      if (!userId) {
        this.logger.warn('Could not determine user ID from client state');
        return { success: false, message: 'Invalid client state format' };
      }

      // Determine the type of change (created, updated, deleted)
      let eventType: string | null;
      switch (changeType) {
        case 'created':
          eventType = OutlookEventTypes.EVENT_CREATED;
          break;
        case 'updated':
          eventType = OutlookEventTypes.EVENT_UPDATED;
          break;
        case 'deleted':
          eventType = OutlookEventTypes.EVENT_DELETED;
          break;
        default:
          eventType = null;
          this.logger.warn(`Unknown change type received: ${String(changeType)}`);
          return { success: false, message: `Unsupported change type: ${String(changeType)}` };
      }

      // Process the resource data
      const resourceData: OutlookResourceData = {
        id: '',
        userId,
        subscriptionId,
        resource,
        changeType,
      };

      // Emit an event for other parts of the application to handle
      if (eventType) {
        this.eventEmitter.emit(eventType, resourceData);
        this.logger.log(`Processed webhook notification: ${eventType}`);
      }

      return { success: true, message: 'Notification processed' };
    } catch (error: unknown) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Error processing webhook notification: ${errorMessage}`);
      return { success: false, message: errorMessage };
    }
  }
}