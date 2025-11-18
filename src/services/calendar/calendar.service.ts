import { Injectable, Logger, Inject, forwardRef } from "@nestjs/common";
import { EventEmitter2 } from "@nestjs/event-emitter";
import { Client } from "@microsoft/microsoft-graph-client";
import axios from "axios";
import { Event, Calendar, Subscription, ChangeNotification } from "../../types";
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
import { DeltaSyncService, DeltaEvent } from "../shared/delta-sync.service";
import { delay, retryWithBackoff } from "../../utils/retry.util";

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
    private readonly deltaSyncService: DeltaSyncService
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
        await this.microsoftAuthService.getUserAccessTokenByExternalUserId(
          externalUserId
        );

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
    try {
      // Get a valid access token for this user
      const accessToken =
        await this.microsoftAuthService.getUserAccessTokenByExternalUserId(
          externalUserId
        );

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
    }
  }

  async deleteEvent(
    event: Partial<Event>,
    externalUserId: string,
    calendarId: string
  ): Promise<void> {
    try {
      // Get a valid access token for this user
      const accessToken =
        await this.microsoftAuthService.getUserAccessTokenByExternalUserId(
          externalUserId
        );

      // Initialize Microsoft Graph client
      const client = Client.init({
        authProvider: (done) => {
          done(null, accessToken);
        },
      });
      this.logger.log(`Deleting event ${event.id} from calendar ${calendarId} for user ${externalUserId}`);
      // Delete the event
      (await client
        .api(`/me/calendars/${calendarId}/events/${event.id}`)
        .delete()) as Event;
    } catch (error: unknown) {
      const errorMessage =
        error instanceof Error ? error.message : "Unknown error";
      this.logger.error(
        `Failed to delete Outlook calendar event: ${errorMessage}`
      );
      throw new Error(
        `Failed to delete Outlook calendar event: ${errorMessage}`
      );
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
      const accessToken =
        await this.microsoftAuthService.getUserAccessTokenByExternalUserId(
          externalUserId
        );

      // Set expiration date (max 3 days as per Microsoft documentation)
      const expirationDateTime = new Date();
      expirationDateTime.setHours(expirationDateTime.getHours() + 72); // 3 days from now

      const appUrl =
        this.microsoftConfig.backendBaseUrl || "http://localhost:3000";
      const basePath = this.microsoftConfig.basePath;
      const basePathUrl = basePath ? `${appUrl}/${basePath}` : appUrl;

      // Create subscription payload with proper URL encoding
      const notificationUrl = `${basePathUrl}/calendar/webhook`;

      // Create subscription payload
      const subscriptionData = {
        changeType: "created,updated,deleted",
        notificationUrl,
        // Add lifecycleNotificationUrl for increased reliability
        lifecycleNotificationUrl: notificationUrl,
        resource: "/me/events",
        expirationDateTime: expirationDateTime.toISOString(),
        clientState: `user_${externalUserId}_${Math.random().toString(36).substring(2, 15)}`,
      };

      this.logger.debug(
        `Creating webhook subscription with notificationUrl: ${notificationUrl}`
      );

      this.logger.debug(
        `Subscription data: ${JSON.stringify(subscriptionData)}`
      );
      // Create the subscription with Microsoft Graph API
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
        `Created webhook subscription ${response.data.id || "unknown"} for user ${externalUserId}`
      );

      // Store internal userId for webhooks (should be the numeric ID in our subscription table)
      const internalUserId = parseInt(externalUserId, 10);

      // Save the subscription to the database
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

      this.logger.debug(`Stored subscription`);

      return response.data;
    } catch (error) {
      this.logger.error("Failed to create webhook subscription:", error);
      throw new Error("Failed to create webhook subscription");
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
      const accessToken =
        await this.microsoftAuthService.getUserAccessTokenByExternalUserId(
          externalUserId
        );

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
            "Content-Type": "application/json",
          },
        }
      );

      // Update the expiration date in our database
      if (response.data.expirationDateTime) {
        await this.webhookSubscriptionRepository.updateSubscriptionExpiration(
          subscriptionId,
          new Date(response.data.expirationDateTime)
        );
      }

      this.logger.log(`Renewed webhook subscription: ${subscriptionId}`);

      return response.data;
    } catch (error: unknown) {
      const errorMessage =
        error instanceof Error ? error.message : "Unknown error";
      this.logger.error(
        `Failed to renew webhook subscription: ${errorMessage}`
      );
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
      const accessToken =
        await this.microsoftAuthService.getUserAccessTokenByUserId(
          internalUserId
        );

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
            "Content-Type": "application/json",
          },
        }
      );

      // Update the expiration date in our database
      if (response.data.expirationDateTime) {
        await this.webhookSubscriptionRepository.updateSubscriptionExpiration(
          subscriptionId,
          new Date(response.data.expirationDateTime)
        );
      }

      this.logger.log(`Renewed webhook subscription: ${subscriptionId}`);

      return response.data;
    } catch (error: unknown) {
      const errorMessage =
        error instanceof Error ? error.message : "Unknown error";
      this.logger.error(
        `Failed to renew webhook subscription: ${errorMessage}`
      );
      throw new Error(`Failed to renew webhook subscription: ${errorMessage}`);
    }
  }

  /**
   * Delete a webhook subscription
   * @param subscriptionId - ID of the subscription to delete
   * @param externalUserId - External user ID for the subscription
   * @returns True if deletion was successful
   */
  async deleteWebhookSubscription(
    subscriptionId: string,
    externalUserId: string
  ): Promise<boolean> {
    try {
      // Get a valid access token for this user
      const accessToken =
        await this.microsoftAuthService.getUserAccessTokenByExternalUserId(
          externalUserId
        );

      // Make the request to Microsoft Graph API to delete the subscription
      await axios.delete(
        `https://graph.microsoft.com/v1.0/subscriptions/${subscriptionId}`,
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
        }
      );

      // Remove the subscription from our database
      await this.webhookSubscriptionRepository.deactivateSubscription(
        subscriptionId
      );

      await this.microsoftUserRepository.update({ externalUserId }, {
        isActive: false
      });

      this.logger.log(`Deleted webhook subscription: ${subscriptionId}`);

      return true;
    } catch (error: unknown) {
      const errorMessage =
        error instanceof Error ? error.message : "Unknown error";
      this.logger.error(
        `Failed to delete webhook subscription: ${errorMessage}`
      );

      // If we get a 404, the subscription doesn't exist anymore at Microsoft,
      // so we should remove it from our database
      if (axios.isAxiosError(error) && error.response?.status === 404) {
        await this.webhookSubscriptionRepository.deactivateSubscription(
          subscriptionId
        );
        this.logger.log(
          `Subscription not found, removed from database: ${subscriptionId}`
        );
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
          // Renew the subscription using the userId to get a fresh token
          await this.renewWebhookSubscriptionByUserId(
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
   * @returns Success status and message
   */
  async handleOutlookWebhook(
    notificationItem: ChangeNotification
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

      // External user Id is the client application userId
      const externalUserId = subscription.userId;

      if (!externalUserId) {
        this.logger.warn(
          "Could not determine external user ID from client state"
        );
        return { success: false, message: "Invalid client state format" };
      }

      const sortedChanges = await this.fetchAndSortChanges(
        String(externalUserId)
      );

      // Process each change and emit appropriate events
      for (const change of sortedChanges) {
        let eventType: string | null;

        // If the change has the @removed property, it's a deletion
        if ((change as { ["@removed"]?: unknown })["@removed"]) {
          eventType = OutlookEventTypes.EVENT_DELETED;
        } else {
          console.log(
            change.createdDateTime,
            change.lastModifiedDateTime,
            change.subject
          );
          eventType =
            !change.createdDateTime ||
            new Date(
              change.lastModifiedDateTime ?? change.createdDateTime
            ).getTime() -
              new Date(change.createdDateTime).getTime() <=
              1000
              ? // If lastModifiedDateTime - createdDateTime is less than a second, it's a new even
                OutlookEventTypes.EVENT_CREATED
              : // Otherwise, it's an update
                OutlookEventTypes.EVENT_UPDATED;
        }

        const resourceData: OutlookResourceData = {
          id: change.id || "",
          userId: externalUserId,
          subscriptionId,
          resource,
          changeType:
            eventType === "outlook.event.deleted"
              ? "deleted"
              : eventType === "outlook.event.created"
                ? "created"
                : "updated",
          data: change as unknown as Record<string, unknown>,
        };

        // Emit the event
        this.eventEmitter.emit(eventType, resourceData);
        this.logger.log(
          `Processed calendar change: ${eventType} for event ID: ${change.id || "unknown"}`
        );
      }

      return { success: true, message: "Notification processed" };
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
   * @returns Array of events sorted by lastModifiedDateTime
   */
  async fetchAndSortChanges(externalUserId: string): Promise<Event[]> {
    const client = await this.getAuthenticatedClient(externalUserId);
    const requestUrl = "/me/events/delta";

    try {
      const events =
        await this.deltaSyncService.fetchAndSortChanges<DeltaEvent>(
          client,
          requestUrl,
          externalUserId
        );

      return events as Event[];
    } catch (error) {
      this.logger.error("Error fetching delta changes:", error);
      throw error;
    }
  }

  async getAuthenticatedClient(externalUserId: string): Promise<Client> {
    const accessToken =
      await this.microsoftAuthService.getUserAccessTokenByExternalUserId(
        externalUserId
      );

    return Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
    });
  }

  async getEventDetails(
    resource: string,
    externalUserId: string
  ): Promise<Event | null> {
    try {
      // Get a valid access token for this user
      const accessToken =
        await this.microsoftAuthService.getUserAccessTokenByExternalUserId(
          externalUserId
        );

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
      this.logger.error("Error fetching event details:", error);
      throw error;
    }
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
