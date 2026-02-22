import { Injectable, Logger, Inject, forwardRef } from "@nestjs/common";
import { EventEmitter2 } from "@nestjs/event-emitter";
import { Client } from "@microsoft/microsoft-graph-client";
import axios from "axios";
import { Event, Calendar, Subscription, ChangeNotification, BatchRequestPayload, BatchResponsePayload } from "../../types";
import { MicrosoftAuthService } from "../auth/microsoft-auth.service";
import { Cron, CronExpression } from "@nestjs/schedule";
import { OutlookWebhookSubscriptionRepository } from "../../repositories/outlook-webhook-subscription.repository";
import { OutlookDeltaLinkRepository } from "../../repositories/outlook-delta-link.repository";
import { OutlookResourceData } from "../../dto/outlook-webhook-notification.dto";
import { MICROSOFT_CONFIG } from "../../constants";
import { MicrosoftOutlookConfig } from "../../interfaces/config/outlook-config.interface";
import { OutlookEventTypes } from "../../enums/event-types.enum";
import { InjectRepository } from "@nestjs/typeorm";
import { MicrosoftUser } from "../../entities/microsoft-user.entity";
import { Repository } from "typeorm";
import { DeltaSyncService } from "../shared/delta-sync.service";
import { delay, retryWithBackoff, is404Error } from "../../utils/retry.util";
import { UserIdConverterService } from "../shared/user-id-converter.service";
import { ResourceType } from "../../enums/resource-type.enum";
import { MicrosoftSubscriptionService } from "../subscription/microsoft-subscription.service";
import { executeGraphApiCall } from "../../utils/outlook-api-executor.util";
import { OutlookWebhookSubscription } from "../../entities/outlook-webhook-subscription.entity";
import { GraphRateLimiterService } from "../shared/graph-rate-limiter.service";

// Event type constants
const OUTLOOK_EVENT_CREATED = OutlookEventTypes.EVENT_CREATED;
const OUTLOOK_EVENT_UPDATED = OutlookEventTypes.EVENT_UPDATED;
const OUTLOOK_EVENT_DELETED = OutlookEventTypes.EVENT_DELETED;

// Change type mapping
const EVENT_TYPE_TO_CHANGE_TYPE: Record<string, "created" | "updated" | "deleted"> = {
  [OUTLOOK_EVENT_CREATED]: "created",
  [OUTLOOK_EVENT_UPDATED]: "updated",
  [OUTLOOK_EVENT_DELETED]: "deleted",
};

/**
 * Check if an event is newly created based on timestamps
 * An event is considered new if lastModifiedDateTime - createdDateTime <= 1 second
 */
function isNewEvent(change: Event): boolean {
  if (!change.createdDateTime) {
    return true;
  }

  const lastModified = new Date(
    change.lastModifiedDateTime ?? change.createdDateTime
  ).getTime();
  const created = new Date(change.createdDateTime).getTime();

  return lastModified - created <= 1000;
}

/**
 * Detect the event type based on change properties
 */
function detectEventType(change: Event): OutlookEventTypes {
  // Check if the change represents a deletion
  if ((change as { ["@removed"]?: unknown })["@removed"]) {
    return OUTLOOK_EVENT_DELETED;
  }

  // Determine if it's a creation or update based on timestamps
  return isNewEvent(change) ? OUTLOOK_EVENT_CREATED : OUTLOOK_EVENT_UPDATED;
}

@Injectable()
export class CalendarService {
  private readonly logger = new Logger(CalendarService.name);
  private readonly syncLocks = new Map<string, Promise<void>>();

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
    private readonly userIdConverter: UserIdConverterService,
    private readonly subscriptionService: MicrosoftSubscriptionService,
    private readonly rateLimiter: GraphRateLimiterService
  ) {}

  /**
   * Get the user's default calendar ID
   * @param externalUserId - External user ID
   * @returns The default calendar ID
   */
  async getDefaultCalendarId(externalUserId: string): Promise<string> {
    try {
      // Get a valid access token for this user
      const accessToken =
        await this.microsoftAuthService.getUserAccessToken({externalUserId});

      // Using axios for direct API call
      const response = await axios.get<Calendar>(
        "https://graph.microsoft.com/v1.0/me/calendar",
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
        }
      );

      if (!response.data.id) {
        throw new Error("Failed to retrieve calendar ID");
      }

      return response.data.id;
    } catch (error) {
      this.logger.error("Error getting default calendar ID:", error);
      throw new Error("Failed to get calendar ID from Microsoft");
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
    calendarId: string
  ): Promise<{ event: Event }> {
    // Rate limit before making request
    await this.rateLimiter.acquirePermit(externalUserId);

    try {
      // Get a valid access token for this user
      const accessToken =
        await this.microsoftAuthService.getUserAccessToken({externalUserId});

      // Initialize Microsoft Graph client
      const client = Client.init({
        authProvider: (done) => {
          done(null, accessToken);
        },
      });

      // Create the event
      const createdEvent = (await client
        .api(`/me/calendars/${calendarId}/events`)
        .post(event)) as Event;

      // Return just the event
      return {
        event: createdEvent,
      };
    } catch (error: unknown) {
      const errorMessage =
        error instanceof Error ? error.message : "Unknown error";
      this.logger.error(
        `Failed to create Outlook calendar event: ${errorMessage}`
      );
      throw new Error(
        `Failed to create Outlook calendar event: ${errorMessage}`
      );
    } finally {
      this.rateLimiter.releasePermit(externalUserId);
    }
  }

  /**
   * Updates an existing event in the user's Outlook calendar
   * @param eventId - The ID of the event to update
   * @param updates - Partial Event object with fields to update
   * @param externalUserId - External user ID associated with the calendar
   * @param calendarId - Calendar ID where the event exists
   * @returns The updated event data
   */
  async updateEvent(
    eventId: string,
    updates: Partial<Event>,
    externalUserId: string,
    calendarId: string
  ): Promise<{ event: Event }> {
    // Rate limit before making request
    await this.rateLimiter.acquirePermit(externalUserId);

    try {
      // Get a valid access token for this user
      const accessToken =
        await this.microsoftAuthService.getUserAccessToken({externalUserId});

      // Initialize Microsoft Graph client
      const client = Client.init({
        authProvider: (done) => {
          done(null, accessToken);
        },
      });

      this.logger.log(`Updating event ${eventId} in calendar ${calendarId} for user ${externalUserId}`);

      // PATCH the existing event
      const updatedEvent = (await client
        .api(`/me/calendars/${calendarId}/events/${eventId}`)
        .patch(updates)) as Event;

      return {
        event: updatedEvent,
      };
    } catch (error: unknown) {
      const errorMessage =
        error instanceof Error ? error.message : "Unknown error";
      this.logger.error(
        `Failed to update Outlook calendar event: ${errorMessage}`
      );
      throw new Error(
        `Failed to update Outlook calendar event: ${errorMessage}`
      );
    } finally {
      this.rateLimiter.releasePermit(externalUserId);
    }
  }

  async deleteEvent(
    event: Partial<Event>,
    externalUserId: string,
    calendarId: string
  ): Promise<void> {
    // Rate limit before making request
    await this.rateLimiter.acquirePermit(externalUserId);

    try {
      // Get a valid access token for this user
      const internalUserId = await this.userIdConverter.toInternalUserId(externalUserId);
      const accessToken =
        await this.microsoftAuthService.getUserAccessToken({internalUserId});

      // Initialize Microsoft Graph client
      const client = Client.init({
        authProvider: (done) => {
          done(null, accessToken);
        },
      });
      this.logger.log(`Deleting event ${event.id} from calendar ${calendarId} for user ${externalUserId}`);

      // Delete the event with retry logic for transient failures
      await retryWithBackoff(
        async () => {
          await client
            .api(`/me/calendars/${calendarId}/events/${event.id}`)
            .delete();
        },
        { maxRetries: 3, retryDelayMs: 1000 }
      );
    } catch (error: unknown) {
      // If the event doesn't exist (404), deletion intent is fulfilled — treat as success
      if (is404Error(error)) {
        this.logger.warn(
          `Outlook calendar event ${event.id} not found (already deleted), treating as successful deletion`
        );
        return;
      }

      const errorMessage =
        error instanceof Error ? error.message : "Unknown error";
      this.logger.error(
        `Failed to delete Outlook calendar event after retries: ${errorMessage}`
      );
      throw new Error(
        `Failed to delete Outlook calendar event: ${errorMessage}`
      );
    } finally {
      this.rateLimiter.releasePermit(externalUserId);
    }
  }

  /**
   * Delete multiple events in a single batch request
   * Uses Microsoft Graph $batch API for efficient batch deletion
   *
   * @param eventIds - Array of event IDs to delete
   * @param externalUserId - External user ID
   * @param calendarId - Calendar ID
   * @returns Results array with success/failure for each event
   *
   * @remarks
   * - Processes up to 20 events per batch (Microsoft Graph limit)
   * - Returns success for 404 errors (already deleted events)
   * - Continues processing even if some deletions fail
   */
  /**
   * Create multiple events in a single batch request
   * Uses Microsoft Graph $batch API for efficient batch creation
   *
   * @param events - Array of event objects to create
   * @param externalUserId - External user ID
   * @param calendarId - Calendar ID
   * @returns Results array with success/failure for each event
   *
   * @remarks
   * - Processes up to 20 events per batch (Microsoft Graph limit)
   * - Returns created event data for successful creations
   * - Continues processing even if some creations fail
   */
  async createBatchEvents(
    events: Partial<Event>[],
    externalUserId: string,
    calendarId: string
  ): Promise<{ index: number; success: boolean; event?: Event; error?: string }[]> {
    if (events.length === 0) {
      return [];
    }

    try {
      const internalUserId = await this.userIdConverter.toInternalUserId(externalUserId);
      const accessToken = await this.microsoftAuthService.getUserAccessToken({ internalUserId });

      const results: { index: number; success: boolean; event?: Event; error?: string }[] = [];

      // Microsoft Graph batch API has a limit of 20 requests per batch
      const BATCH_SIZE = 20;

      for (let i = 0; i < events.length; i += BATCH_SIZE) {
        const batchEvents = events.slice(i, i + BATCH_SIZE);

        // Rate limit before making batch request
        await this.rateLimiter.acquirePermit(externalUserId);

        try {
          // Build batch request payload
          const batchPayload: BatchRequestPayload = {
            requests: batchEvents.map((event, index) => ({
              id: `${index}`,
              method: 'POST',
              url: `/me/calendars/${calendarId}/events`,
              body: event,
              headers: {
                'Content-Type': 'application/json',
              },
            })),
          };

          this.logger.log(
            `Creating batch of ${batchEvents.length} events in calendar ${calendarId} for user ${externalUserId}`
          );

          // Execute batch request with retry logic
          const batchResponse = await retryWithBackoff(
            async () => {
              const response = await axios.post<BatchResponsePayload<Event>>(
                'https://graph.microsoft.com/v1.0/$batch',
                batchPayload,
                {
                  headers: {
                    Authorization: `Bearer ${accessToken}`,
                    'Content-Type': 'application/json',
                  },
                }
              );
              return response.data;
            },
            { maxRetries: 3, retryDelayMs: 1000 }
          );

          // Process batch results
          batchResponse.responses.forEach((response, batchIndex) => {
            const globalIndex = i + batchIndex;

            // Success: 201 (Created)
            if (response.status === 201) {
              results.push({
                index: globalIndex,
                success: true,
                event: response.body,
              });

              this.logger.debug(`Successfully created event at index ${globalIndex}`);
            } else {
              // Failure: any other status code
              const errorMessage = JSON.stringify(response.body);
              results.push({
                index: globalIndex,
                success: false,
                error: `HTTP ${response.status}: ${errorMessage}`,
              });

              this.logger.error(`Failed to create event at index ${globalIndex}: ${errorMessage}`);
            }
          });

        } catch (error) {
          // Batch request failed entirely - mark all events in this batch as failed
          const errorMessage = error instanceof Error ? error.message : 'Unknown error';

          batchEvents.forEach((_, batchIndex) => {
            const globalIndex = i + batchIndex;
            results.push({
              index: globalIndex,
              success: false,
              error: `Batch request failed: ${errorMessage}`,
            });
          });

          this.logger.error(
            `Batch creation failed for ${batchEvents.length} events: ${errorMessage}`
          );
        } finally {
          this.rateLimiter.releasePermit(externalUserId);
        }
      }

      const successCount = results.filter(r => r.success).length;
      const failCount = results.filter(r => !r.success).length;

      this.logger.log(
        `Batch creation complete: ${successCount} succeeded, ${failCount} failed out of ${events.length} total`
      );

      return results;

    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to execute batch creation: ${errorMessage}`);

      // Return failure for all events
      return events.map((_, index) => ({
        index,
        success: false,
        error: `Batch creation failed: ${errorMessage}`,
      }));
    }
  }

  /**
   * Update multiple events in a single batch request
   * Uses Microsoft Graph $batch API for efficient batch updates
   *
   * @param updates - Array of update objects with eventId and fields to update
   * @param externalUserId - External user ID
   * @param calendarId - Calendar ID
   * @returns Results array with success/failure for each update
   *
   * @remarks
   * - Processes up to 20 events per batch (Microsoft Graph limit)
   * - Returns updated event data for successful updates
   * - Continues processing even if some updates fail
   */
  async updateBatchEvents(
    updates: Array<{ eventId: string; updates: Partial<Event> }>,
    externalUserId: string,
    calendarId: string
  ): Promise<{ index: number; success: boolean; event?: Event; error?: string }[]> {
    if (updates.length === 0) {
      return [];
    }

    try {
      const internalUserId = await this.userIdConverter.toInternalUserId(externalUserId);
      const accessToken = await this.microsoftAuthService.getUserAccessToken({ internalUserId });

      const results: { index: number; success: boolean; event?: Event; error?: string }[] = [];

      // Microsoft Graph batch API has a limit of 20 requests per batch
      const BATCH_SIZE = 20;

      for (let i = 0; i < updates.length; i += BATCH_SIZE) {
        const batchUpdates = updates.slice(i, i + BATCH_SIZE);

        // Build batch request payload
        const batchPayload: BatchRequestPayload = {
          requests: batchUpdates.map((update, index) => ({
            id: `${index}`,
            method: 'PATCH',
            url: `/me/calendars/${calendarId}/events/${update.eventId}`,
            body: update.updates,
            headers: {
              'Content-Type': 'application/json',
            },
          })),
        };

        this.logger.log(
          `Updating batch of ${batchUpdates.length} events in calendar ${calendarId} for user ${externalUserId}`
        );

        try {
          // Execute batch request with retry logic
          const batchResponse = await retryWithBackoff(
            async () => {
              const response = await axios.post<BatchResponsePayload<Event>>(
                'https://graph.microsoft.com/v1.0/$batch',
                batchPayload,
                {
                  headers: {
                    Authorization: `Bearer ${accessToken}`,
                    'Content-Type': 'application/json',
                  },
                }
              );
              return response.data;
            },
            { maxRetries: 3, retryDelayMs: 1000 }
          );

          // Process batch results
          batchResponse.responses.forEach((response, batchIndex) => {
            const globalIndex = i + batchIndex;

            // Success: 200 (OK)
            if (response.status === 200) {
              results.push({
                index: globalIndex,
                success: true,
                event: response.body,
              });

              this.logger.debug(`Successfully updated event at index ${globalIndex}`);
            } else {
              // Failure: any other status code
              const errorMessage = JSON.stringify(response.body);
              results.push({
                index: globalIndex,
                success: false,
                error: `HTTP ${response.status}: ${errorMessage}`,
              });

              this.logger.error(`Failed to update event at index ${globalIndex}: ${errorMessage}`);
            }
          });

        } catch (error) {
          // Batch request failed entirely - mark all events in this batch as failed
          const errorMessage = error instanceof Error ? error.message : 'Unknown error';

          batchUpdates.forEach((_, batchIndex) => {
            const globalIndex = i + batchIndex;
            results.push({
              index: globalIndex,
              success: false,
              error: `Batch request failed: ${errorMessage}`,
            });
          });

          this.logger.error(
            `Batch update failed for ${batchUpdates.length} events: ${errorMessage}`
          );
        }
      }

      const successCount = results.filter(r => r.success).length;
      const failCount = results.filter(r => !r.success).length;

      this.logger.log(
        `Batch update complete: ${successCount} succeeded, ${failCount} failed out of ${updates.length} total`
      );

      return results;

    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to execute batch update: ${errorMessage}`);

      // Return failure for all events
      return updates.map((_, index) => ({
        index,
        success: false,
        error: `Batch update failed: ${errorMessage}`,
      }));
    }
  }

  async deleteBatchEvents(
    eventIds: string[],
    externalUserId: string,
    calendarId: string
  ): Promise<{ id: string; success: boolean; error?: string }[]> {
    if (eventIds.length === 0) {
      return [];
    }

    try {
      const internalUserId = await this.userIdConverter.toInternalUserId(externalUserId);
      const accessToken = await this.microsoftAuthService.getUserAccessToken({ internalUserId });

      const results: { id: string; success: boolean; error?: string }[] = [];

      // Microsoft Graph batch API has a limit of 20 requests per batch
      const BATCH_SIZE = 20;

      for (let i = 0; i < eventIds.length; i += BATCH_SIZE) {
        const batchEventIds = eventIds.slice(i, i + BATCH_SIZE);

        // Build batch request payload
        const batchPayload: BatchRequestPayload = {
          requests: batchEventIds.map((eventId, index) => ({
            id: `${index}`,
            method: 'DELETE',
            url: `/me/calendars/${calendarId}/events/${eventId}`,
          })),
        };

        this.logger.log(
          `Deleting batch of ${batchEventIds.length} events from calendar ${calendarId} for user ${externalUserId}`
        );

        try {
          // Execute batch request with retry logic
          const batchResponse = await retryWithBackoff(
            async () => {
              const response = await axios.post<BatchResponsePayload>(
                'https://graph.microsoft.com/v1.0/$batch',
                batchPayload,
                {
                  headers: {
                    Authorization: `Bearer ${accessToken}`,
                    'Content-Type': 'application/json',
                  },
                }
              );
              return response.data;
            },
            { maxRetries: 3, retryDelayMs: 1000 }
          );

          // Process batch results
          batchResponse.responses.forEach((response, index) => {
            const eventId = batchEventIds[index];

            // Success: 204 (No Content) or 404 (already deleted)
            if (response.status === 204 || response.status === 404) {
              results.push({ id: eventId, success: true });

              if (response.status === 404) {
                this.logger.warn(
                  `Event ${eventId} not found (already deleted), treating as successful deletion`
                );
              }
            } else {
              // Failure: any other status code
              const errorMessage = response.body ? JSON.stringify(response.body) : 'Unknown error';
              results.push({
                id: eventId,
                success: false,
                error: `HTTP ${response.status}: ${errorMessage}`,
              });

              this.logger.error(`Failed to delete event ${eventId}: ${errorMessage}`);
            }
          });

        } catch (error) {
          // Batch request failed entirely - mark all events in this batch as failed
          const errorMessage = error instanceof Error ? error.message : 'Unknown error';

          batchEventIds.forEach(eventId => {
            results.push({
              id: eventId,
              success: false,
              error: `Batch request failed: ${errorMessage}`,
            });
          });

          this.logger.error(
            `Batch deletion failed for ${batchEventIds.length} events: ${errorMessage}`
          );
        }
      }

      const successCount = results.filter(r => r.success).length;
      const failCount = results.filter(r => !r.success).length;

      this.logger.log(
        `Batch deletion complete: ${successCount} succeeded, ${failCount} failed out of ${eventIds.length} total`
      );

      return results;

    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to execute batch deletion: ${errorMessage}`);

      // Return failure for all events
      return eventIds.map(id => ({
        id,
        success: false,
        error: `Batch deletion failed: ${errorMessage}`,
      }));
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
        clientState: `user_${internalUserId}_${Math.random().toString(36).substring(2, 15)}`,
      };

      this.logger.log(
        `[${correlationId}] Creating webhook subscription with notificationUrl: ${notificationUrl}`
      );

      this.logger.debug(
        `[${correlationId}] Subscription data: ${JSON.stringify(subscriptionData)}`
      );
      // Create the subscription with Microsoft Graph API
      this.logger.log(`[${correlationId}] Sending POST request to Microsoft Graph API`);
      const response = await axios.post<Subscription>(
        "https://graph.microsoft.com/v1.0/subscriptions",
        subscriptionData,
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
        }
      );

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
   *
   * @example
   * ```typescript
   * const renewed = await calendarService.renewWebhookSubscription(
   *   'sub-456-xyz',
   *   102  // Internal user ID from database
   * );
   * ```
   */
  async renewWebhookSubscription(
    subscriptionId: string,
    internalUserId: number
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

      // Make the request to Microsoft Graph API to renew the subscription
      const response = await axios.patch<Subscription>(
        `https://graph.microsoft.com/v1.0/subscriptions/${subscriptionId}`,
        renewalData,
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
        }
      );

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

        // Subscription no longer exists at Microsoft
        if (statusCode === 404) {
          this.logger.warn(
            `[${correlationId}] Subscription ${subscriptionId} not found at Microsoft. ` +
            `Deactivating in local database.`
          );

          await this.webhookSubscriptionRepository.deactivateSubscription(
            subscriptionId
          );

          throw new Error(
            `Subscription ${subscriptionId} not found at Microsoft. ` +
            `Subscription has been automatically deactivated.`
          );
        }

        // User token issues (401, 403)
        if (statusCode === 401 || statusCode === 403) {
          this.logger.error(
            `[${correlationId}] Authentication failed for subscription ${subscriptionId}. ` +
            `Status: ${statusCode}, Response: ${JSON.stringify(error.response?.data)}`
          );
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

  async getSubscription(subscriptionId: string): Promise<OutlookWebhookSubscription | null> {
    const subscription = await this.webhookSubscriptionRepository.findBySubscriptionId(subscriptionId);

    if (!subscription) {
      return null;
    }

    return subscription;
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
   *
   * @example
   * // Using external ID (common in public API)
   * await calendarService.deleteWebhookSubscription('sub-456', 'user-7');
   *
   * // Using internal ID (common in cleanup flows)
   * await calendarService.deleteWebhookSubscription('sub-456', 102);
   */
  async deleteWebhookSubscription(
    subscriptionId: string,
    userId: string | number
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

      // Delegate to MicrosoftSubscriptionService for subscription deletion
      this.logger.debug(
        `[${correlationId}] Calling MicrosoftSubscriptionService to delete subscription`
      );

      this.logger.log(`Deleted webhook subscription: ${subscriptionId}`);
      await this.subscriptionService.deleteSubscription(subscriptionId, accessToken);

      this.logger.log(
        `[${correlationId}] Successfully deleted subscription at Microsoft`
      );

      // Remove the subscription from our database (soft delete)
      await this.webhookSubscriptionRepository.deactivateSubscription(
        subscriptionId
      );

      this.logger.debug(
        `[${correlationId}] Deactivated subscription in local database`
      );

      // Mark user as inactive
      const updateCriteria = typeof userId === 'string' ? { externalUserId: userId } : { id: userId };
      await this.microsoftUserRepository.update(
        updateCriteria,
        { isActive: false }
      );

      this.logger.debug(
        `[${correlationId}] Marked Microsoft user as inactive (${typeof userId === 'string' ? 'externalUserId' : 'id'}: ${userId})`
      );

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

  /**
   * Get active webhook subscription for a user
   * @param externalUserId - External user ID from host application
   * @returns Subscription ID if active subscription exists, null otherwise
   */
  async getActiveSubscription(externalUserId: string): Promise<string | null> {
    try {
      // Convert external to internal ID
      const internalUserId = await this.userIdConverter.externalToInternal(externalUserId);

      this.logger.log(`[getActiveSubscription] Getting active subscription for user ${externalUserId} (internalUserId: ${internalUserId})`);
      const subscription = await this.webhookSubscriptionRepository.findActiveByUserId(internalUserId);
      this.logger.log(`[getActiveSubscription] Found subscription: ${subscription?.subscriptionId}`);
      return subscription?.subscriptionId ?? null;
    } catch {
      // User may not have connected Microsoft account yet - this is not an error
      this.logger.debug(`No active subscription for user ${externalUserId}`);
      return null;
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
      const expiringSubscriptions =
        await this.webhookSubscriptionRepository.findSubscriptionsNeedingRenewal(
          24 // hours until expiration
        );

      if (expiringSubscriptions.length === 0) {
        this.logger.debug("No subscriptions need renewal");
        return;
      }

      this.logger.log(
        `Found ${String(expiringSubscriptions.length)} subscriptions that need renewal`
      );

      // Renew each subscription
      for (const subscription of expiringSubscriptions) {
        try {
          // Renew the subscription using the internal userId to get a fresh token
          await this.renewWebhookSubscription(
            subscription.subscriptionId,
            subscription.userId
          );
        } catch (error) {
          this.logger.error(
            `Failed to renew subscription ${subscription.subscriptionId}:`,
            error
          );
          // Continue with the next subscription even if this one failed
        }
      }
    } catch (error) {
      this.logger.error("Error in subscription renewal job:", error);
    }
  }

  /**
   * Handle a webhook notification from Microsoft
   * @param notificationItem - The notification data from Microsoft
   * @param useStreaming - Whether to use streaming mode (default: false for buffering)
   * @returns Success status and message
   */
  async handleOutlookWebhook(
    notificationItem: ChangeNotification,
    useStreaming: boolean = false
  ): Promise<{ success: boolean; message: string }> {
    try {
      // Extract necessary information from the notification
      const { subscriptionId, clientState, resource, changeType } =
        notificationItem;

      this.logger.debug(
        `Received webhook notification for subscription: ${subscriptionId || "unknown"}`
      );
      this.logger.debug(
        `Resource: ${resource || "unknown"}, ChangeType: ${String(changeType || "unknown")}`
      );

      // Validate subscription and extract user ID
      const {success, internalUserId, message} = await this.validateWebhookSubscription(
        subscriptionId,
        clientState
      );

      if (!success || !internalUserId) {
        this.logger.error('validateWebhookSubscription failed', message || 'Unknown error');
        return { success: false, message: message || 'Unknown error' };
      }

      // Process changes using appropriate strategy (passed as parameter)
      const totalProcessed = useStreaming
        ? await this.processChangesStreaming(
            internalUserId,
            String(subscriptionId || ''),
            resource || ''
          )
        : await this.processChangesBuffering(
            internalUserId,
            String(subscriptionId || ''),
            resource || ''
          );

      return { success: true, message: `Processed ${totalProcessed} events` };
    } catch (error: unknown) {
      const errorMessage =
        error instanceof Error ? error.message : "Unknown error";
      this.logger.error(
        `Error processing webhook notification: ${errorMessage}`
      );
      return { success: false, message: errorMessage };
    }
  }

  /**
   * Fetches and sorts calendar changes using delta API
   * @param externalUserId External user ID
   * @param forceReset Force reset delta link (used on reconnection)
   * @param dateRange Optional date range for calendar delta queries (only used on initialization)
   * @returns Array of events sorted by lastModifiedDateTime
   */
  async fetchAndSortChanges(
    externalUserId: string,
    forceReset: boolean = false,
    dateRange?: {
      startDate: Date;
      endDate: Date;
    }
  ): Promise<Event[]> {
    const client = await this.getAuthenticatedClient(externalUserId);
    const requestUrl = "/me/events/delta";

    try {
      this.logger.log(`[fetchAndSortChanges] Starting delta fetch for user ${externalUserId} (forceReset: ${forceReset})`);

      const items = await this.deltaSyncService.fetchAndSortChanges(
        client,
        requestUrl,
        externalUserId,
        forceReset,
        dateRange
      );

      this.logger.log(`[fetchAndSortChanges] ✅ Successfully fetched ${items.length} calendar changes for user ${externalUserId}`);
      return items as Event[];
    } catch (error) {
      // Enhanced error logging with context
      const errorDetails = this.extractErrorDetails(error);

      this.logger.error(
        `[fetchAndSortChanges] ❌ Failed to fetch delta changes for user ${externalUserId}`,
        {
          userId: externalUserId,
          forceReset,
          errorType: errorDetails.type,
          statusCode: errorDetails.statusCode,
          errorCode: errorDetails.code,
          errorMessage: errorDetails.message,
          timestamp: new Date().toISOString(),
        }
      );

      // Log full error stack for debugging
      this.logger.error("Error fetching delta changes:", error);
      throw error;
    }
  }

  /**
   * Extract detailed error information from Microsoft Graph errors
   * @param error - The error object
   * @returns Structured error details
   */
  private extractErrorDetails(error: unknown): {
    type: string;
    statusCode: number | string;
    code: string;
    message: string;
  } {
    if (!error || typeof error !== 'object') {
      return {
        type: 'unknown',
        statusCode: 'N/A',
        code: 'N/A',
        message: String(error),
      };
    }

    // Check for Microsoft Graph SDK error format
    if ('statusCode' in error) {
      const statusCode = error.statusCode as number;
      const code = 'code' in error ? String(error.code) : 'N/A';
      const body = 'body' in error ? String(error.body) : 'N/A';

      // Detect network errors
      if (statusCode === -1) {
        return {
          type: 'network_error',
          statusCode: -1,
          code: code,
          message: 'Network connectivity failure - unable to reach Microsoft Graph API. Check internet connection, firewall, or proxy settings.',
        };
      }

      return {
        type: 'graph_api_error',
        statusCode: statusCode,
        code: code,
        message: body,
      };
    }

    // Check for nested error in stack array
    if ('stack' in error && Array.isArray(error.stack) && error.stack.length > 0) {
      const firstError: unknown = error.stack[0];
      if (firstError && typeof firstError === 'object') {
        const statusCode: number | string = 'statusCode' in firstError ? (firstError.statusCode as number) : 'N/A';
        const code: string = 'code' in firstError ? String(firstError.code) : 'N/A';
        const body: string = 'body' in firstError ? String(firstError.body) : 'N/A';

        if (typeof statusCode === 'number' && statusCode === -1) {
          return {
            type: 'network_error',
            statusCode: -1,
            code: code,
            message: 'Network connectivity failure - unable to reach Microsoft Graph API',
          };
        }

        return {
          type: 'graph_api_error',
          statusCode: statusCode,
          code: code,
          message: body,
        };
      }
    }

    // Generic error
    return {
      type: 'generic_error',
      statusCode: 'N/A',
      code: 'N/A',
      message: error instanceof Error ? error.message : JSON.stringify(error),
    };
  }

  /**
   * Streams calendar changes using async generator (memory-efficient alternative)
   * Yields sorted batches of events as each page is fetched from Microsoft Graph
   *
   * Benefits:
   * - Memory efficient: Process one page at a time instead of loading all changes
   * - Faster time-to-first-event: Start processing immediately after first page
   * - Better for large syncs: Handle 1000s of changes without memory issues
   *
   * @param externalUserId External user ID
   * @param forceReset Force reset delta link (used on reconnection)
   * @param dateRange Optional date range for calendar delta queries (only used on initialization)
   * @param saveDeltaLink Whether to save the delta link to database (default: true). Set to false for one-time windowed imports.
   * @yields Sorted batches of events (one batch per Microsoft Graph page)
   * @returns Final delta link (null if not available)
   */
  async *streamCalendarChanges(
    externalUserId: string,
    forceReset: boolean = false,
    dateRange?: {
      startDate: Date;
      endDate: Date;
    },
    saveDeltaLink: boolean = true
  ): AsyncGenerator<Event[], string | null, unknown> {
    const client = await this.getAuthenticatedClient(externalUserId);
    const requestUrl = "/me/events/delta";

    try {
      this.logger.log(`[streamCalendarChanges] Starting stream for user ${externalUserId} (saveDeltaLink: ${saveDeltaLink})`);

      const deltaLink = yield* this.deltaSyncService.streamDeltaChanges(
        client,
        requestUrl,
        externalUserId,
        forceReset,
        dateRange,
        saveDeltaLink
      );

      this.logger.log(`[streamCalendarChanges] Completed streaming for user ${externalUserId}, deltaLink: ${deltaLink ? 'received' : 'none'}`);

      return deltaLink;
    } catch (error) {
      this.logger.error(`[streamCalendarChanges] Error streaming delta changes:`, error instanceof Error ? error.message : String(error));
      throw error;
    }
  }

  async getAuthenticatedClient(externalUserId: string): Promise<Client> {
    const accessToken =
      await this.microsoftAuthService.getUserAccessToken({externalUserId});

    return Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
    });
  }

  /**
   * Save delta link for calendar resource
   * @param externalUserId External user ID
   * @param deltaLink Delta link to save
   */
  async saveDeltaLink(externalUserId: string, deltaLink: string): Promise<void> {
    const internalUserId = await this.userIdConverter.externalToInternal(externalUserId);
    await this.deltaLinkRepository.saveDeltaLink(internalUserId, ResourceType.CALENDAR, deltaLink);
    this.logger.log(`[saveDeltaLink] Saved delta link for user ${externalUserId} (internal: ${internalUserId})`);
  }

  async getEventDetails(
    resource: string,
    externalUserId: string
  ): Promise<Event | null> {
    try {
      // Get a valid access token for this user
      const accessToken =
        await this.microsoftAuthService.getUserAccessToken({externalUserId});

      const response = await axios.get(
        `https://graph.microsoft.com/v1.0/${resource}`,
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
          },
        }
      );

      return response.data as Event;
    } catch (error) {
      // If event deleted (404), return null gracefully
      if (axios.isAxiosError(error) && error.response?.status === 404) {
        this.logger.debug(
          `Event not found (404) for resource ${resource}, likely deleted - returning null`
        );
        return null;
      }

      // For other errors, throw
      this.logger.error("Error fetching event details:", error);
      throw error;
    }
  }

  /**
   * Fetch multiple events in a single batch request using Microsoft Graph JSON Batch API
   *
   * This method uses the Microsoft Graph /$batch endpoint to fetch up to 20 events
   * in a single HTTP request, significantly improving performance over individual calls.
   *
   * Enhanced with:
   * - Per-user rate limiting to prevent 429 errors
   * - Automatic retry of individual 429s within batch responses
   * - Retry-After header support for cooldown periods
   *
   * @param eventIds - Array of event IDs to fetch (max 20 per Microsoft Graph limit)
   * @param externalUserId - External user ID
   * @returns Array of successfully fetched events
   * @throws Error if batch request fails or access token cannot be obtained
   *
   * @example
   * const events = await calendarService.getEventsBatch(
   *   ['event-id-1', 'event-id-2', 'event-id-3'],
   *   'user-123'
   * );
   *
   * @remarks
   * - Maximum 20 events per batch (Microsoft Graph limit)
   * - Handles partial failures gracefully (404s are logged, not thrown)
   * - Retries individual 429s up to 2 times per event
   * - Uses per-user rate limiting (4 req/sec, 10k req/10min)
   */
  async getEventsBatch(
    eventIds: string[],
    externalUserId: string
  ): Promise<Event[]> {
    return this.getEventsBatchInternal(eventIds, externalUserId, new Map());
  }

  /**
   * Internal implementation of getEventsBatch with retry tracking
   *
   * @param eventIds - Event IDs to fetch
   * @param externalUserId - External user ID
   * @param retryCount - Map tracking retry attempts per event ID
   * @param maxRetries - Maximum retries per event (default: 2)
   * @returns Successfully fetched events
   * @private
   */
  private async getEventsBatchInternal(
    eventIds: string[],
    externalUserId: string,
    retryCount: Map<string, number>,
    maxRetries = 2
  ): Promise<Event[]> {
    if (eventIds.length === 0) {
      return [];
    }

    if (eventIds.length > 20) {
      this.logger.warn(
        `[getEventsBatch] Called with ${eventIds.length} events, exceeding limit. Only first 20 will be fetched.`
      );
    }

    // Rate limit BEFORE making batch request
    await this.rateLimiter.acquirePermit(externalUserId);

    try {
      // Get a valid access token for this user
      const accessToken =
        await this.microsoftAuthService.getUserAccessToken({externalUserId});

      // Build batch request payload (max 20 requests)
      const batchPayload: BatchRequestPayload = {
        requests: eventIds.slice(0, 20).map((eventId, index) => ({
          id: `${index}`,
          method: 'GET',
          url: `/me/events/${eventId}`,
        })),
      };

      this.logger.debug(
        `[getEventsBatch] Fetching ${batchPayload.requests.length} events in batch for user ${externalUserId}`
      );

      // Execute batch request with automatic retry and rate limit handling
      const response = await executeGraphApiCall(
        () => axios.post<BatchResponsePayload<Event>>(
          'https://graph.microsoft.com/v1.0/$batch',
          batchPayload,
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
              'Content-Type': 'application/json',
            },
          }
        ),
        {
          maxRetries: 3,
          retryDelayMs: 1000,
          logger: this.logger,
          resourceName: `batch events (user: ${externalUserId})`,
        }
      );

      // Handle null response (shouldn't happen for batch requests, but type safety)
      if (!response) {
        this.logger.error(`[getEventsBatch] Batch request returned null for user ${externalUserId}`);
        return [];
      }

      // Process responses
      const events: Event[] = [];
      const rateLimitedEventIds: string[] = [];
      let successCount = 0;
      let notFoundCount = 0;
      let errorCount = 0;

      for (const batchResponse of response.data.responses) {
        const eventId = eventIds[parseInt(batchResponse.id, 10)];

        if (batchResponse.status === 200) {
          events.push(batchResponse.body);
          successCount++;
        } else if (batchResponse.status === 404) {
          // Event was deleted between drift detection and fetching
          this.logger.warn(
            `[getEventsBatch] Event ${eventId} not found (404), likely deleted`
          );
          notFoundCount++;
        } else if (batchResponse.status === 429) {
          // Individual event rate limited - queue for retry
          rateLimitedEventIds.push(eventId);
          this.logger.warn(
            `[getEventsBatch] Event ${eventId} rate limited (429) within batch response`
          );
          errorCount++;
        } else {
          // Other errors - log and continue
          this.logger.error(
            `[getEventsBatch] Event ${eventId} failed: status ${batchResponse.status}, ` +
            `body: ${JSON.stringify(batchResponse.body)}`
          );
          errorCount++;
        }
      }

      this.logger.log(
        `[getEventsBatch] Batch complete for user ${externalUserId}: ` +
        `success=${successCount}, notFound=${notFoundCount}, errors=${errorCount}`
      );

      // Retry rate-limited events (with per-event retry limit)
      if (rateLimitedEventIds.length > 0) {
        const retryableEvents = rateLimitedEventIds.filter(id => {
          const count = retryCount.get(id) || 0;
          return count < maxRetries;
        });

        if (retryableEvents.length > 0) {
          // Extract Retry-After from response headers if present
          const retryAfter = this.extractRetryAfterFromResponse(response);
          if (retryAfter) {
            this.rateLimiter.handleRateLimitResponse(externalUserId, retryAfter);
          }

          // Update retry counts
          retryableEvents.forEach(id => {
            retryCount.set(id, (retryCount.get(id) || 0) + 1);
          });

          const maxAttempt = Math.max(...retryableEvents.map(id => retryCount.get(id) || 0));
          this.logger.log(
            `[getEventsBatch] Retrying ${retryableEvents.length} rate-limited events ` +
            `(attempt ${maxAttempt}/${maxRetries})`
          );

          // Recursive retry with updated counts
          const retriedEvents = await this.getEventsBatchInternal(
            retryableEvents,
            externalUserId,
            retryCount,
            maxRetries
          );

          events.push(...retriedEvents);
        } else {
          this.logger.error(
            `[getEventsBatch] ${rateLimitedEventIds.length} events exceeded max retries - DATA LOSS WARNING`
          );
        }
      }

      return events;
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(
        `[getEventsBatch] Batch request failed for user ${externalUserId}: ${errorMessage}`
      );
      throw error;
    } finally {
      // Always release permit
      this.rateLimiter.releasePermit(externalUserId);
    }
  }

  /**
   * Extract Retry-After header from batch response
   * @param response - Axios response from batch request
   * @returns Retry-After seconds, or null if not present
   * @private
   */
  private extractRetryAfterFromResponse(
    response: { headers?: Record<string, unknown> }
  ): number | null {
    try {
      const retryAfter = response.headers?.['retry-after'];
      if (typeof retryAfter === 'string' || typeof retryAfter === 'number') {
        const parsed = parseInt(String(retryAfter), 10);
        return !isNaN(parsed) ? Math.max(parsed, 5) : null;
      }
    } catch (_error) {
      this.logger.debug('[extractRetryAfterFromResponse] Failed to extract Retry-After header');
    }
    return null;
  }

  /**
   * Stream calendar events in chunks for memory efficiency
   *
   * This method uses an async generator pattern to stream events in configurable batch sizes,
   * minimizing memory usage for large calendars. Events are fetched from Microsoft Graph API
   * using the calendarView endpoint with pagination and automatic retry logic.
   *
   * @param externalUserId - External user ID
   * @param options - Optional configuration
   * @param options.startDate - Optional start date filter (defaults to today)
   * @param options.endDate - Optional end date filter (defaults to 5 years from now)
   * @param options.batchSize - Number of events to yield per chunk (default 100)
   * @yields Chunks of events (Event[])
   * @throws Error if authentication fails or API requests fail after retries
   *
   * @example
   * // Basic usage - stream all events with default settings
   * for await (const events of calendarService.importEventsStream('user-123')) {
   *   console.log(`Processing ${events.length} events`);
   *   // Process events in batches of 100
   * }
   *
   * @example
   * // Stream events with custom date range
   * const startDate = new Date('2024-01-01');
   * const endDate = new Date('2024-12-31');
   *
   * for await (const events of calendarService.importEventsStream('user-123', {
   *   startDate,
   *   endDate,
   *   batchSize: 50
   * })) {
   *   // Process 2024 events in batches of 50
   *   await saveEventsToDatabase(events);
   * }
   *
   * @example
   * // Collect all events (memory-intensive for large calendars)
   * const allEvents: Event[] = [];
   * for await (const chunk of calendarService.importEventsStream('user-123')) {
   *   allEvents.push(...chunk);
   * }
   * console.log(`Total events: ${allEvents.length}`);
   *
   * @example
   * // Stream with progress tracking
   * let totalProcessed = 0;
   * const stream = calendarService.importEventsStream('user-123', { batchSize: 200 });
   *
   * for await (const events of stream) {
   *   totalProcessed += events.length;
   *   console.log(`Progress: ${totalProcessed} events processed`);
   *
   *   // Process events with custom logic
   *   for (const event of events) {
   *     await processEvent(event);
   *   }
   * }
   *
   * @remarks
   * - Memory footprint: ~1MB constant vs 10-30MB for loading all events
   * - Automatic exponential backoff retry (3 attempts) on API failures
   * - 200ms delay between API pages to respect rate limits
   * - Emits IMPORT_COMPLETED event when streaming finishes
   * - Uses calendarView endpoint which automatically expands recurring events
   */
  async *importEventsStream(
    externalUserId: string,
    options?: {
      startDate?: Date;
      endDate?: Date;
      batchSize?: number;
    }
  ): AsyncGenerator<Event[], void, unknown> {
    const batchSize = options?.batchSize ?? 100;

    try {
      this.logger.log(
        `Starting event stream for user ${externalUserId} (batchSize: ${batchSize})`
      );

      const client = await this.getAuthenticatedClient(externalUserId);

      // Build request URL
      const requestUrl = this.buildRequestUrl(options, batchSize);

      let nextLink: string | undefined = requestUrl;
      const buffer: Event[] = [];
      let totalFetched = 0;

      // Fetch pages and yield chunks
      while (nextLink) {
        this.logger.debug(`Fetching page: ${nextLink}`);

        // Fetch with retry logic
        const response = (await retryWithBackoff(() =>
          client.api(nextLink as string).get()
        )) as { value: Event[]; "@odata.nextLink"?: string };

        const items: Event[] = response.value;
        buffer.push(...items);
        totalFetched += items.length;

        // Yield when buffer reaches batch size
        while (buffer.length >= batchSize) {
          const chunk = buffer.splice(0, batchSize);
          this.logger.debug(
            `Yielding chunk of ${chunk.length} items (total fetched: ${totalFetched})`
          );
          yield chunk;
        }

        nextLink = response["@odata.nextLink"];

        // Small delay between pages for rate limiting
        if (nextLink) {
          await delay(200);
        }
      }

      // Yield remaining items in buffer
      if (buffer.length > 0) {
        this.logger.debug(`Yielding final chunk of ${buffer.length} items`);
        yield buffer;
      }

      this.logger.log(
        `Completed streaming ${totalFetched} events for user ${externalUserId}`
      );

      // Emit completion event with metadata
      this.eventEmitter.emit(OutlookEventTypes.IMPORT_COMPLETED, {
        userId: externalUserId,
        totalEvents: totalFetched,
      });
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : "Unknown error";
      this.logger.error(
        `Error streaming events for user ${externalUserId}: ${errorMessage}`
      );
      throw error;
    }
  }

  async handleOutlookWebhookV2(
    notificationItem: ChangeNotification,
  ) {
    if (!notificationItem.subscriptionId) {
      this.logger.error(`Subscription ID is required`);
      return { success: false, message: `Subscription ID is required` };
    }

    try {
      const subscription = await this.getSubscription(notificationItem.subscriptionId);

      if (!subscription) {
        this.logger.error(`Subscription not found for subscriptionId: ${notificationItem.subscriptionId}`);
        return { success: false, message: `Subscription not found for subscriptionId: ${notificationItem.subscriptionId}` };
      }

      this.eventEmitter.emit(OutlookEventTypes.EVENT_NOTIFICATION, {
        userId: subscription.userId,
        subscriptionId: subscription.subscriptionId,
        clientState: subscription.clientState,
        resource: notificationItem,
        changeType: notificationItem.changeType,
      });

      return { success: true, message: 'Notification processed' };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Error handling Outlook webhook v2: ${errorMessage}`);
      return { success: false, message: errorMessage };
    }
  }

  /**
   * Stream instances of a specific recurring event series
   *
   * Uses the Microsoft Graph /me/events/{seriesMasterId}/instances endpoint to fetch
   * only the expanded occurrences of a specific recurring event within a date range.
   * This is more targeted than calendarView which returns ALL events.
   *
   * @param seriesMasterId - The ID of the recurring event series master
   * @param externalUserId - External user ID
   * @param options - Optional date range and batch size
   * @yields Batches of Event instances for the recurring series
   */
  async *getRecurringEventInstances(
    seriesMasterId: string,
    externalUserId: string,
    options?: { startDate?: Date; endDate?: Date; batchSize?: number }
  ): AsyncGenerator<Event[], void, unknown> {
    const batchSize = options?.batchSize ?? 100;

    try {
      const client = await this.getAuthenticatedClient(externalUserId);

      const now = new Date();
      const startDate = options?.startDate ?? new Date(now.getFullYear(), now.getMonth() - 1, now.getDate());
      const endDate = options?.endDate ?? new Date(now.getFullYear(), now.getMonth() + 6, now.getDate());

      // Microsoft Graph API's endDateTime parameter is EXCLUSIVE (only returns occurrences starting BEFORE this time).
      // To include occurrences that happen ON the end date, we add 1 day to make it inclusive.
      // Example: For a weekly Tuesday series ending Jan 28, without this adjustment, the Jan 28 occurrence
      // (starting at 10:00 AM) would be excluded because it doesn't start before midnight Jan 28.
      const inclusiveEndDate = new Date(endDate);
      inclusiveEndDate.setDate(inclusiveEndDate.getDate() + 1);

      this.logger.log(
        `[getRecurringEventInstances] Fetching instances for series ${seriesMasterId} from ${startDate.toISOString()} to ${endDate.toISOString()} (API endDateTime: ${inclusiveEndDate.toISOString()})`
      );

      let nextLink: string | undefined =
        `/me/events/${seriesMasterId}/instances?startDateTime=${startDate.toISOString()}&endDateTime=${inclusiveEndDate.toISOString()}&$top=${batchSize}`;

      const buffer: Event[] = [];
      let totalFetched = 0;

      while (nextLink) {
        const response = (await retryWithBackoff(() =>
          client.api(nextLink as string).get()
        )) as { value: Event[]; '@odata.nextLink'?: string };

        const items: Event[] = response.value;
        buffer.push(...items);
        totalFetched += items.length;

        while (buffer.length >= batchSize) {
          const chunk = buffer.splice(0, batchSize);
          yield chunk;
        }

        nextLink = response['@odata.nextLink'];
        if (nextLink) {
          await delay(200);
        }
      }

      if (buffer.length > 0) {
        yield buffer;
      }

      this.logger.log(
        `[getRecurringEventInstances] Completed: ${totalFetched} instances for series ${seriesMasterId}`
      );
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(
        `[getRecurringEventInstances] Error fetching instances for series ${seriesMasterId}: ${errorMessage}`
      );
      throw error;
    }
  }

  /**
   * Initialize delta sync tracking without importing events
   *
   * Call this AFTER manual import to establish baseline for incremental sync.
   * This method initializes the delta link, allowing you to track ALL future
   * calendar changes regardless of date range.
   *
   * Use case:
   * 1. Import events in a specific date range (e.g., next 3 months) using importEventsStream
   * 2. Call this method to enable tracking of ALL future changes (not limited to that range)
   *
   * Note: This method fetches all current events from Microsoft Graph to establish
   * the delta baseline, but the events are intentionally not returned. Use
   * importEventsStream for initial import, then call this to enable webhooks.
   *
   * @param externalUserId - External user ID
   *
   * @example
   * await calendarService.initializeDeltaSync(userId);
   * → Enables tracking of ALL future calendar changes (not limited to a window range)
   */
  async initializeDeltaSync(externalUserId: string): Promise<void> {
    this.logger.log(`Initializing delta sync tracking for user ${externalUserId}`);

    try {
      const client = await this.getAuthenticatedClient(externalUserId);

      // Convert external ID to internal ID
      const internalUserId = await this.userIdConverter.externalToInternal(externalUserId);

      // Initialize delta link WITHOUT date range = tracks ALL events going forward
      // Events returned are intentionally ignored - we only need the delta token
      await this.deltaSyncService.initializeDeltaLink(
        client,
        "/me/events/delta",
        internalUserId,
        ResourceType.CALENDAR
      );

      this.logger.log(`Delta tracking enabled for user ${externalUserId} (all events)`);
    } catch (error) {
      this.logger.error(
        `Failed to initialize delta sync for user ${externalUserId}: ${
          error instanceof Error ? error.message : "Unknown error"
        }`
      );
      throw error;
    }
  }

  /**
   * Process delta changes using streaming mode (page-by-page)
   * Lower memory footprint, processes each page immediately as it arrives
   *
   * @param externalUserId External user ID
   * @param subscriptionId Subscription ID
   * @param resource Resource path
   * @returns Number of events processed
   */
  private async processChangesStreaming(
    internalUserId: number,
    subscriptionId: string,
    resource: string
  ): Promise<number> {
    let totalProcessed = 0;
    let batchCount = 0;

    const externalUserId = await this.userIdConverter.internalToExternal(internalUserId);
    this.logger.log(`[processChangesStreaming] Using STREAMING mode for user ${internalUserId}`);

    for await (const changeBatch of this.streamCalendarChanges(externalUserId)) {
      batchCount++;
      this.logger.log(`[processChangesStreaming] Processing batch ${batchCount} with ${changeBatch.length} changes`);

      if (changeBatch.length === 0) {
        this.logger.warn(`[processChangesStreaming] Received empty batch ${batchCount}`);
        continue;
      }

      // Process each change in this batch
      for (const change of changeBatch) {
        this.processDeltaEventChange(
          change,
          internalUserId,
          subscriptionId,
          resource
        );
        totalProcessed++;
      }

      this.logger.log(`[processChangesStreaming] Batch ${batchCount} processed: ${changeBatch.length} events (total: ${totalProcessed})`);
    }

    this.logger.log(`[processChangesStreaming] Completed: ${totalProcessed} events across ${batchCount} batches`);
    return totalProcessed;
  }

  /**
   * Process delta changes using buffering mode (fetch all, then process)
   * Higher memory usage but fewer backend calls
   *
   * @param externalUserId External user ID
   * @param subscriptionId Subscription ID
   * @param resource Resource path
   * @returns Number of events processed
   */
  private async processChangesBuffering(
    internalUserId: number,
    subscriptionId: string,
    resource: string
  ): Promise<number> {
    const externalUserId = await this.userIdConverter.internalToExternal(internalUserId);

    this.logger.log(`[processChangesBuffering] Using BUFFERING mode for user ${internalUserId}`);

    const allChanges = await this.fetchAndSortChanges(externalUserId);

    if (allChanges.length === 0) {
      this.logger.warn(`[processChangesBuffering] No changes found`);
      return 0;
    }

    this.logger.log(`[processChangesBuffering] Fetched ${allChanges.length} changes, processing batch`);

    let totalProcessed = 0;

    // Process all changes
    for (const change of allChanges) {
      this.processDeltaEventChange(
        change,
        internalUserId,
        subscriptionId,
        resource
      );
      totalProcessed++;
    }

    this.logger.log(`[processChangesBuffering] Completed: ${totalProcessed} events processed`);
    return totalProcessed;
  }

  /**
   * Process a single delta event change and emit appropriate event
   * @param change The delta event change to process
   * @param externalUserId External user ID
   * @param subscriptionId Subscription ID
   * @param resource Resource path
   */
  private processDeltaEventChange(
    change: Event,
    internalUserId: number,
    subscriptionId: string,
    resource: string
  ): void {
    const eventType = detectEventType(change);

    this.logger.debug(
      `[processDeltaEventChange] Event ${change.id || "unknown"}: created=${change.createdDateTime}, modified=${change.lastModifiedDateTime}, type=${eventType}`
    );

    const resourceData: OutlookResourceData = {
      id: change.id || "",
      userId: internalUserId,
      subscriptionId,
      resource,
      changeType: EVENT_TYPE_TO_CHANGE_TYPE[eventType],
      data: change as unknown as Record<string, unknown>,
    };

    // Emit the event
    this.eventEmitter.emit(eventType, resourceData);

    this.logger.log(
      `[processDeltaEventChange] Emitted ${eventType} for event ID: ${change.id || "unknown"}`
    );
  }

  /**
   * Validate webhook subscription and extract user ID
   * @param subscriptionId - Subscription ID from notification
   * @param clientState - Client state from notification
   * @returns Validation result with user ID or error
   */
  private async validateWebhookSubscription(
    subscriptionId: string | undefined,
    clientState: string | null | undefined
  ): Promise<{ success: boolean; internalUserId?: number; message?: string }> {
    // Find the subscription in our database to verify it's legitimate
    const subscription =
      await this.webhookSubscriptionRepository.findBySubscriptionId(
        subscriptionId || ""
      );

    if (!subscription) {
      this.logger.warn(
        `Unknown subscription ID: ${subscriptionId || "unknown"}`
      );
      return { success: false, message: "Unknown subscription" };
    }

    // Verify the client state for additional security
    if (
      subscription.clientState &&
      clientState !== subscription.clientState
    ) {
      this.logger.warn("Client state mismatch");
      return { success: false, message: "Client state mismatch" };
    }

    const internalUserId = subscription.userId;

    if (!internalUserId) {
      this.logger.warn(
        "Could not determine internal user ID from client state"
      );
      return { success: false, message: "Invalid client state format" };
    }

    return { success: true, internalUserId };
  }

  /**
   * Build request URL for event import
   * @param options - Import options
   * @param batchSize - Batch size for pagination
   * @returns Request URL
   */
  private buildRequestUrl(
    options?: {
      startDate?: Date;
      endDate?: Date;
    },
    batchSize?: number
  ): string {
    // Build the request URL for full import with date range
    // Use calendarView for proper date range queries and recurring event handling
    // Microsoft Graph API limits calendarView to max 1825 days (5 years) total range
    const dateinterval = 5 * 365 * 24 * 60 * 60 * 1000; // 5 years
    const defaultStartDate = new Date(); // Today
    const defaultEndDate = new Date(Date.now() + dateinterval); // 5 years from now
    const startDate = options?.startDate ?? defaultStartDate;
    const endDate = options?.endDate ?? defaultEndDate;

    const startDateStr = startDate.toISOString();
    const endDateStr = endDate.toISOString();

    // Build base URL with required parameters
    let url = `/me/calendarView?startDateTime=${startDateStr}&endDateTime=${endDateStr}`;

    // Add ordering and pagination
    url += `&$orderby=start/dateTime&$top=${batchSize ?? 100}`;

    return url;
  }
}
