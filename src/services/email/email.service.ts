import { Injectable, Logger, Inject, forwardRef } from '@nestjs/common';
import { EventEmitter2 } from '@nestjs/event-emitter';
import { Client } from '@microsoft/microsoft-graph-client';
import axios from 'axios';
import { MicrosoftAuthService } from '../auth/microsoft-auth.service';
import { MICROSOFT_CONFIG } from '../../constants';
import { MicrosoftOutlookConfig } from '../../interfaces/config/outlook-config.interface';
import { Message, ChangeNotification, Subscription } from '@microsoft/microsoft-graph-types';
import { OutlookWebhookSubscriptionRepository } from '../../repositories/outlook-webhook-subscription.repository';
import { OutlookResourceData } from '../../dto/outlook-webhook-notification.dto';
import { OutlookEventTypes } from '../../enums/event-types.enum';
import { MicrosoftUser } from '../../entities/microsoft-user.entity';
import { InjectRepository } from '@nestjs/typeorm';
import { Repository, Not } from 'typeorm';
import { UserIdConverterService } from '../shared/user-id-converter.service';
import { executeGraphApiCall } from '../../utils/outlook-api-executor.util';

@Injectable()
export class EmailService {
  private readonly logger = new Logger(EmailService.name);

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
  ) {}

  /**
   * Sends an email using Microsoft Graph API
   * 
   * @param message - The email message to send
   * @param externalUserId - External user ID associated with the email account
   * @returns The sent message data
   */
  async sendEmail(
    message: Partial<Message>,
    externalUserId: string,
  ): Promise<{ message: Message }> {
    try {
      // Get a valid access token for this user
      const accessToken = await this.microsoftAuthService.getUserAccessToken({externalUserId});

      // Initialize Microsoft Graph client
      const client = Client.init({
        authProvider: (done) => {
          done(null, accessToken);
        },
      });

      // Send the email with retry and rate limiting
      const sentMessage = await executeGraphApiCall(
        () => client.api('/me/sendMail').post({ message }),
        {
          logger: this.logger,
          resourceName: `send email for user ${externalUserId}`,
          maxRetries: 3,
        }
      ) as Message;

      // Return just the message
      return {
        message: sentMessage,
      };
    } catch (error: unknown) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to send email: ${errorMessage}`);
      throw new Error(`Failed to send email: ${errorMessage}`);
    }
  }

  /**
   * Create a webhook subscription to receive notifications for incoming emails
   * @param externalUserId - External user ID
   * @returns The created subscription data
   */
  async createWebhookSubscription(
    externalUserId: string,
  ): Promise<Subscription> {
    try {
      // Convert external user ID to internal database ID
      const internalUserId = await this.userIdConverter.externalToInternal(externalUserId);

      // Get a valid access token for this user
      const accessToken = await this.microsoftAuthService.getUserAccessToken({internalUserId});
      
      // Set expiration date (max 3 days as per Microsoft documentation)
      const expirationDateTime = new Date();
      expirationDateTime.setHours(expirationDateTime.getHours() + 72); // 3 days from now

      const appUrl = this.microsoftConfig.backendBaseUrl || 'http://localhost:3000';
      const basePath = this.microsoftConfig.basePath;
      const basePathUrl = basePath ? `${appUrl}/${basePath}` : appUrl;

      // Create subscription payload with proper URL encoding
      const notificationUrl = `${basePathUrl}/email/webhook`;

      // Create subscription payload
      const subscriptionData = {
        changeType: 'created,updated,deleted',
        notificationUrl,
        // Add lifecycleNotificationUrl for increased reliability
        lifecycleNotificationUrl: notificationUrl,
        resource: '/me/messages',
        expirationDateTime: expirationDateTime.toISOString(),
        clientState: `user_${internalUserId}_${Math.random().toString(36).substring(2, 15)}`,
      };

      this.logger.debug(`Creating email webhook subscription with notificationUrl: ${notificationUrl}`);

      this.logger.debug(`Subscription data: ${JSON.stringify(subscriptionData)}`);

      // Create the subscription with Microsoft Graph API with retry
      const response = await executeGraphApiCall(
        () => axios.post<Subscription>(
          'https://graph.microsoft.com/v1.0/subscriptions',
          subscriptionData,
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
              'Content-Type': 'application/json',
            },
          }
        ),
        {
          logger: this.logger,
          resourceName: `create email webhook subscription for user ${internalUserId}`,
          maxRetries: 3,
        }
      );

      if (!response || !response.data) {
        throw new Error('Email subscription creation returned null response');
      }

      this.logger.log(`Created email webhook subscription ${response.data.id || 'unknown'} for user ${internalUserId}`);

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
    } catch (error: unknown) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to create email webhook subscription: ${errorMessage}`);
      throw new Error(`Failed to create email webhook subscription: ${errorMessage}`);
    }
  }

  /**
   * Delete an email webhook subscription
   *
   * Deletes the subscription at Microsoft Graph API and deactivates it locally.
   * Identical implementation to CalendarService for consistency.
   *
   * @param subscriptionId - ID of the subscription to delete at Microsoft
   * @param userId - User ID (can be external string or internal number)
   * @returns True if deletion was successful
   * @throws Error if user not found or deletion fails (except 404)
   *
   * @example
   * await emailService.deleteWebhookSubscription('sub-789', 'user-7');
   */
  async deleteWebhookSubscription(
    subscriptionId: string,
    userId: string | number
  ): Promise<boolean> {
    const correlationId = `delete-email-sub-${subscriptionId}-${Date.now()}`;

    try {
      this.logger.log(
        `[${correlationId}] Deleting email subscription ${subscriptionId} for user ${userId}`
      );

      const internalUserId = await this.userIdConverter.toInternalUserId(userId);

      // Get access token (including inactive users since we need to clean up their subscriptions)
      const accessToken = await this.microsoftAuthService.getUserAccessToken({
        internalUserId,
        includeInactive: true
      });

      this.logger.debug(`[${correlationId}] Access token obtained`);

      // Make the request to Microsoft Graph API to delete the subscription with retry
      this.logger.debug(
        `[${correlationId}] Calling Microsoft Graph API to delete email subscription`
      );

      await executeGraphApiCall(
        () => axios.delete(
          `https://graph.microsoft.com/v1.0/subscriptions/${subscriptionId}`,
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
              "Content-Type": "application/json",
            },
          }
        ),
        {
          logger: this.logger,
          resourceName: `delete email webhook subscription ${subscriptionId} for user ${internalUserId}`,
          maxRetries: 3,
          return404AsNull: true,
        }
      );

      this.logger.log(
        `[${correlationId}] Successfully deleted email subscription at Microsoft`
      );

      // Remove the subscription from our database (soft delete)
      await this.webhookSubscriptionRepository.deactivateSubscription(
        subscriptionId
      );

      this.logger.debug(
        `[${correlationId}] Deactivated email subscription in local database`
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
        `[${correlationId}] Successfully deleted email subscription ${subscriptionId}`
      );

      return true;

    } catch (error: unknown) {
      const errorMessage = error instanceof Error ? error.message : "Unknown error";

      // If we get a 404, the subscription doesn't exist anymore at Microsoft
      if (axios.isAxiosError(error) && error.response?.status === 404) {
        this.logger.log(
          `[${correlationId}] Email subscription ${subscriptionId} not found at Microsoft, ` +
          `cleaning up local database`
        );

        await this.webhookSubscriptionRepository.deactivateSubscription(
          subscriptionId
        );

        this.logger.log(
          `[${correlationId}] Successfully cleaned up orphaned email subscription ${subscriptionId}`
        );

        return true;
      }

      this.logger.error(
        `[${correlationId}] Failed to delete email subscription ${subscriptionId}: ${errorMessage}`
      );

      throw new Error(`Failed to delete email webhook subscription: ${errorMessage}`);
    }
  }

  /**
   * Handle a webhook notification from Microsoft for email changes
   * @param notificationItem - The notification data from Microsoft
   * @returns Success status and message
   */
  async handleEmailWebhook(
    notificationItem: ChangeNotification,
  ): Promise<{ success: boolean; message: string }> {
    try {
      // Extract necessary information from the notification
      const { subscriptionId, clientState, resource, changeType } = notificationItem;

      this.logger.debug(`Received email webhook notification for subscription: ${subscriptionId || 'unknown'}`);
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

      // Get the internal userId from our subscription
      const internalUserId = subscription.userId;

      if (!internalUserId) {
        this.logger.warn('Could not determine user ID from client state');
        return { success: false, message: 'Invalid client state format' };
      }

      // Determine the type of change (created, updated, deleted)
      let eventType: string | null;
      switch (changeType) {
        case 'created':
          eventType = OutlookEventTypes.EMAIL_RECEIVED;
          break;
        case 'updated':
          eventType = OutlookEventTypes.EMAIL_UPDATED;
          break;
        case 'deleted':
          eventType = OutlookEventTypes.EMAIL_DELETED;
          break;
        default:
          eventType = null;
          this.logger.warn(`Unknown change type received: ${String(changeType)}`);
          return { success: false, message: `Unsupported change type: ${String(changeType)}` };
      }

      // Get additional email data if it's a new email
      let emailData: Record<string, unknown> = {};
      
      if (changeType === 'created' && resource) {
        try {
          // Extract the message ID from the resource path (format: /me/messages/{id})
          const messageId = resource.split('/').pop();
          if (messageId) {
            // Get a valid access token - use internalUserId to get the token
            const accessToken = await this.microsoftAuthService.getUserAccessToken({internalUserId});
            
            // Create a Graph client to fetch the email details
            const client = Client.init({
              authProvider: (done) => {
                done(null, accessToken);
              },
            });
            
            // Get the email message details with retry
            emailData = await executeGraphApiCall(
              () => client
                .api(`/me/messages/${messageId}`)
                .select('id,subject,receivedDateTime,from,toRecipients,ccRecipients,body')
                .get(),
              {
                logger: this.logger,
                resourceName: `email message details for ${messageId}`,
                maxRetries: 3,
              }
            ) as Record<string, unknown>;
              
            this.logger.log(`Retrieved email details for message ID: ${messageId}`);
          }
        } catch (error: unknown) {
          const errorMessage = error instanceof Error ? error.message : 'Unknown error';
          this.logger.error(`Failed to retrieve email details: ${errorMessage}`);
          // Continue processing even if we can't get the email details
        }
      }

      // Process the resource data
      const resourceData: OutlookResourceData = {
        id: '',
        userId: internalUserId,
        subscriptionId,
        resource,
        changeType,
        data: emailData
      };

      // Emit an event for other parts of the application to handle
      if (eventType) {
        this.eventEmitter.emit(eventType, resourceData);
        this.logger.log(`Processed email webhook notification: ${eventType}`);
      }

      return { success: true, message: 'Notification processed' };
    } catch (error: unknown) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Error processing email webhook notification: ${errorMessage}`);
      return { success: false, message: errorMessage };
    }
  }
} 