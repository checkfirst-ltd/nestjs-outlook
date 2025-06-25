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
import { Cron, CronExpression } from '@nestjs/schedule';
import { OutlookWebhookSubscriptionRepository } from '../../repositories/outlook-webhook-subscription.repository';
import { OutlookDeltaLinkRepository } from '../../repositories/outlook-delta-link.repository';
import { OutlookResourceData } from '../../dto/outlook-webhook-notification.dto';
import { MICROSOFT_CONFIG } from '../../constants';
import { MicrosoftOutlookConfig } from '../../interfaces/config/outlook-config.interface';
import { OutlookEventTypes } from '../../enums/event-types.enum';
import { InjectRepository } from '@nestjs/typeorm';
import { MicrosoftUser } from '../../entities/microsoft-user.entity';
import { Repository } from 'typeorm';
import { DeltaSyncService, DeltaEvent } from '../shared/delta-sync.service';
import { getExternalUserIdFromUserId } from '../shared/shared-user.service';

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
    private readonly deltaLinkRepository: OutlookDeltaLinkRepository,
    @InjectRepository(MicrosoftUser)
    private readonly microsoftUserRepository: Repository<MicrosoftUser>,
    private readonly deltaSyncService: DeltaSyncService,
  ) {}

  /**
   * Get the user's default calendar ID
   * @param externalUserId - External user ID
   * @returns The default calendar ID
   */
  async getDefaultCalendarId(externalUserId: string): Promise<string> {
    try {
      // Get a valid access token for this user
      const accessToken = await this.microsoftAuthService.getUserAccessTokenByExternalUserId(externalUserId);
      
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
   * @param externalUserId - External user ID associated with the calendar
   * @param calendarId - Calendar ID where the event will be created
   * @returns The created event data
   */
  async createEvent(
    event: Partial<Event>,
    externalUserId: string,
    calendarId: string,
  ): Promise<{ event: Event }> {
    try {
      // Get a valid access token for this user
      const accessToken = await this.microsoftAuthService.getUserAccessTokenByExternalUserId(externalUserId);
      
      // Initialize Microsoft Graph client
      const client = Client.init({
        authProvider: (done) => {
          done(null, accessToken);
        },
      });

      // Create the event
      const createdEvent = await client
        .api(`/me/calendars/${calendarId}/events`)
        .post(event) as Event;

      // Return just the event
      return {
        event: createdEvent,
      };
    } catch (error: unknown) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to create Outlook calendar event: ${errorMessage}`);
      throw new Error(`Failed to create Outlook calendar event: ${errorMessage}`);
    }
  }

  /**
   * Create a webhook subscription to receive notifications for calendar events
   * @param externalUserId - External user ID
   * @returns The created subscription data
   */
  async createWebhookSubscription(
    externalUserId: string
  ): Promise<Subscription> {
    try {
      // Get a valid access token for this user
      const accessToken = await this.microsoftAuthService.getUserAccessTokenByExternalUserId(externalUserId);
      
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
        clientState: `user_${externalUserId}_${Math.random().toString(36).substring(2, 15)}`,
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

      this.logger.log(`Created webhook subscription ${response.data.id || 'unknown'} for user ${externalUserId}`);

      // Store internal userId for webhooks (should be the numeric ID in our subscription table)
      const internalUserId = parseInt(externalUserId, 10);
      
      // Save the subscription to the database
      await this.webhookSubscriptionRepository.saveSubscription({
        subscriptionId: response.data.id,
        userId: internalUserId,
        resource: response.data.resource,
        changeType: response.data.changeType,
        clientState: response.data.clientState || '',
        notificationUrl: response.data.notificationUrl,
        expirationDateTime: response.data.expirationDateTime ? new Date(response.data.expirationDateTime) : new Date(),
      });

      this.logger.debug(`Stored subscription`);

      return response.data;
    } catch (error) {
      this.logger.error('Failed to create webhook subscription:', error);
      throw new Error('Failed to create webhook subscription');
    }
  }

  /**
   * Renew an existing webhook subscription
   * @param subscriptionId - ID of the subscription to renew
   * @param externalUserId - External user ID for the subscription
   * @returns The renewed subscription data
   */
  async renewWebhookSubscription(
    subscriptionId: string,
    externalUserId: string
  ): Promise<Subscription> {
    try {
      // Get a valid access token for this user
      const accessToken = await this.microsoftAuthService.getUserAccessTokenByExternalUserId(externalUserId);
      
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
   * Renew an existing webhook subscription using internal user ID
   * @param subscriptionId - ID of the subscription to renew
   * @param internalUserId - Internal user ID for the subscription
   * @returns The renewed subscription data
   */
  async renewWebhookSubscriptionByUserId(
    subscriptionId: string,
    internalUserId: number | string
  ): Promise<Subscription> {
    try {
      // Get a valid access token for this user
      const accessToken = await this.microsoftAuthService.getUserAccessTokenByUserId(internalUserId);
      
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
   * @param externalUserId - External user ID for the subscription
   * @returns True if deletion was successful
   */
  async deleteWebhookSubscription(subscriptionId: string, externalUserId: string): Promise<boolean> {
    try {
      // Get a valid access token for this user
      const accessToken = await this.microsoftAuthService.getUserAccessTokenByExternalUserId(externalUserId);
      
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
      const expiringSubscriptions = await this.webhookSubscriptionRepository.findSubscriptionsNeedingRenewal(
        24 // hours until expiration
      );

      if (expiringSubscriptions.length === 0) {
        this.logger.debug('No subscriptions need renewal');
        return;
      }

      this.logger.log(`Found ${String(expiringSubscriptions.length)} subscriptions that need renewal`);

      // Renew each subscription
      for (const subscription of expiringSubscriptions) {
        try {
          // Renew the subscription using the userId to get a fresh token
          await this.renewWebhookSubscriptionByUserId(
            subscription.subscriptionId,
            subscription.userId
          );
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

      const externalUserId = await getExternalUserIdFromUserId(userId, this.microsoftUserRepository, this.logger);
    
      if (!externalUserId) {
        this.logger.warn(`Could not determine externalUserId for user ID ${String(userId)}`);
        return { success: false, message: 'Could not determine external user ID' };
      }
      
      const sortedChanges = await this.fetchAndSortChanges(String(externalUserId));

      // Process each change and emit appropriate events
    for (const change of sortedChanges) {
      let eventType: string | null;
      
      // If the change has the @removed property, it's a deletion
      if ((change as { ['@removed']?: unknown })['@removed']) {
        eventType = OutlookEventTypes.EVENT_DELETED;
      } else if (!change.createdDateTime || 
                new Date(change.createdDateTime).getTime() === 
                new Date((change.lastModifiedDateTime ?? change.createdDateTime)).getTime()) {
        // If createdDateTime equals lastModifiedDateTime, it's a new event
        eventType = OutlookEventTypes.EVENT_CREATED;
      } else {
        // Otherwise, it's an update
        eventType = OutlookEventTypes.EVENT_UPDATED;
      }
      
      const resourceData: OutlookResourceData = {
        id: change.id || '',
        userId,
        subscriptionId,
        resource,
        changeType: eventType === 'outlook.event.deleted' ? 'deleted' : 
                   eventType === 'outlook.event.created' ? 'created' : 'updated',
        data: change as unknown as Record<string, unknown>,
      };
      
      // Emit the event
      this.eventEmitter.emit(eventType, resourceData);
      this.logger.log(`Processed calendar change: ${eventType} for event ID: ${change.id || 'unknown'}`);
    }

      return { success: true, message: 'Notification processed' };
    } catch (error: unknown) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Error processing webhook notification: ${errorMessage}`);
      return { success: false, message: errorMessage };
    }
  }

  /**
   * Fetches and sorts calendar changes using delta API
   * @param externalUserId External user ID
   * @returns Array of events sorted by lastModifiedDateTime
   */
  async fetchAndSortChanges(externalUserId: string): Promise<Event[]> {
    const client = await this.getAuthenticatedClient(externalUserId);
    const requestUrl = '/me/events/delta';

    try {
      const events = await this.deltaSyncService.fetchAndSortChanges<DeltaEvent>(
        client,
        requestUrl
      );

      return events as Event[];
    } catch (error) {
      this.logger.error('Error fetching delta changes:', error);
      throw error;
    }
  }

  async getAuthenticatedClient(externalUserId: string): Promise<Client> {
    const accessToken = await this.microsoftAuthService.getUserAccessTokenByExternalUserId(externalUserId);
    
    return Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
    });
  }
}
