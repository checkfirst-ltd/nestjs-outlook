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
import { Repository } from 'typeorm';
import { UserIdConverterService } from '../shared/user-id-converter.service';

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

      // Send the email
      const sentMessage = await client
        .api('/me/sendMail')
        .post({ message }) as Message;

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
            
            // Get the email message details
            emailData = await client
              .api(`/me/messages/${messageId}`)
              .select('id,subject,receivedDateTime,from,toRecipients,ccRecipients,body')
              .get() as Record<string, unknown>;
              
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