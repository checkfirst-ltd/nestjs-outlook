import { Injectable, Logger, Inject } from '@nestjs/common';
import { EventEmitter2 } from '@nestjs/event-emitter';
import { InjectRepository } from '@nestjs/typeorm';
import { Repository } from 'typeorm';
import axios from 'axios';
import { Event, Calendar, BatchRequestPayload, BatchResponsePayload } from '../../types';
import { MICROSOFT_CONFIG } from '../../constants';
import { MicrosoftOutlookConfig } from '../../interfaces/config/outlook-config.interface';
import { OutlookEventTypes } from '../../enums/event-types.enum';
import { executeGraphApiCall } from '../../utils/outlook-api-executor.util';
import { TtlCache } from '../../utils/ttl-cache.util';
import { AppOnlyAuthService } from '../auth/app-only-auth.service';
import { MicrosoftTenantUser } from '../../entities/microsoft-tenant-user.entity';

/**
 * Service for managing calendars across all users in a tenant using app-only authentication.
 *
 * Unlike CalendarService which uses `/me/*` endpoints with delegated user authentication,
 * TenantCalendarService uses `/users/{microsoftUserId}/*` endpoints with app-only tokens
 * to access any user's calendar within the tenant.
 *
 * Key differences from CalendarService:
 * - Uses app-only (client credentials) authentication instead of user-delegated
 * - Requires Microsoft Graph user ID to identify the target user
 * - No per-user token refresh needed (single tenant-wide token)
 * - Requires admin consent for Application permissions
 *
 * Required Graph API permissions (Application):
 * - Calendars.ReadWrite (read/write all users' calendars)
 *
 * All requests include:
 * - Prefer: IdType="ImmutableId" - for stable event IDs across tenant
 * - Prefer: outlook.timezone="UTC" - for consistent timezone handling
 */
@Injectable()
export class TenantCalendarService {
  private readonly logger = new Logger(TenantCalendarService.name);

  /**
   * Cache for default calendar IDs: key = microsoftUserId, value = calendarId
   * Calendar IDs are immutable per user, so we cache for 1 hour.
   */
  private readonly defaultCalendarIdCache = new TtlCache<string, string>(60 * 60 * 1000);

  /**
   * Standard headers for all Graph API requests
   */
  private readonly standardHeaders = {
    'Content-Type': 'application/json',
    'Prefer': 'IdType="ImmutableId", outlook.timezone="UTC"',
  };

  constructor(
    private readonly eventEmitter: EventEmitter2,
    @Inject(MICROSOFT_CONFIG)
    private readonly microsoftConfig: MicrosoftOutlookConfig,
    private readonly appOnlyAuthService: AppOnlyAuthService,
    @InjectRepository(MicrosoftTenantUser)
    private readonly tenantUserRepository: Repository<MicrosoftTenantUser>,
  ) {}

  /**
   * Build the authorization header with the provided access token
   */
  private buildHeaders(accessToken: string): Record<string, string> {
    return {
      ...this.standardHeaders,
      Authorization: `Bearer ${accessToken}`,
    };
  }

  /**
   * Get the user's default calendar ID.
   *
   * @param tenantId - Microsoft tenant ID (Azure AD tenant GUID)
   * @param microsoftUserId - Microsoft Graph user ID
   * @returns The default calendar ID
   *
   * @example
   * ```typescript
   * const calendarId = await tenantCalendarService.getDefaultCalendarId(
   *   'tenant-guid',
   *   'user-guid-here'
   * );
   * ```
   */
  async getDefaultCalendarId(
    tenantId: string,
    microsoftUserId: string,
  ): Promise<string> {
    // Check cache first
    const cacheKey = `${tenantId}:${microsoftUserId}`;
    const cached = this.defaultCalendarIdCache.get(cacheKey);
    if (cached) {
      this.logger.debug(`[getDefaultCalendarId] Cache hit for user ${microsoftUserId}`);
      return cached;
    }

    this.logger.log(`[getDefaultCalendarId] Fetching calendar ID for user ${microsoftUserId}`);

    try {
      // Get app-only access token for the tenant
      const accessToken = await this.appOnlyAuthService.getAccessToken(tenantId);

      const response = await executeGraphApiCall(
        () => axios.get<Calendar>(
          `https://graph.microsoft.com/v1.0/users/${microsoftUserId}/calendar`,
          {
            headers: this.buildHeaders(accessToken),
          }
        ),
        {
          logger: this.logger,
          resourceName: `users/${microsoftUserId}/calendar`,
          maxRetries: 7,
        }
      );

      if (!response?.data.id) {
        throw new Error(`Failed to retrieve calendar ID for user ${microsoftUserId}`);
      }

      const calendarId = response.data.id;

      // Cache the result
      this.defaultCalendarIdCache.set(cacheKey, calendarId);

      // Also persist to database if we have a tenant user mapping
      await this.updateUserCalendarId(microsoftUserId, calendarId);

      this.logger.log(`[getDefaultCalendarId] Cached calendar ID for user ${microsoftUserId}`);

      return calendarId;
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(
        `[getDefaultCalendarId] Failed to get calendar ID for user ${microsoftUserId}: ${errorMessage}`
      );
      throw new Error(`Failed to get calendar ID from Microsoft: ${errorMessage}`);
    }
  }

  /**
   * Update the cached calendar ID in the database
   */
  private async updateUserCalendarId(
    microsoftUserId: string,
    calendarId: string,
  ): Promise<void> {
    try {
      await this.tenantUserRepository.update(
        { microsoftUserId, isActive: true },
        { defaultCalendarId: calendarId }
      );
    } catch (error) {
      // Non-critical - just log and continue
      this.logger.debug(
        `[updateUserCalendarId] Could not update calendar ID in DB: ${error instanceof Error ? error.message : 'Unknown'}`
      );
    }
  }

  /**
   * Create an event in a user's calendar.
   *
   * @param event - Microsoft Graph Event object with event details
   * @param tenantId - Microsoft tenant ID (Azure AD tenant GUID)
   * @param microsoftUserId - Microsoft Graph user ID
   * @param calendarId - Calendar ID where the event will be created
   * @returns The created event data
   */
  async createEvent(
    event: Partial<Event>,
    tenantId: string,
    microsoftUserId: string,
    calendarId: string,
  ): Promise<{ event: Event }> {
    try {
      this.logger.log(
        `[createEvent] Creating event in calendar ${calendarId} for user ${microsoftUserId}`
      );

      // Get app-only access token for the tenant
      const accessToken = await this.appOnlyAuthService.getAccessToken(tenantId);

      const createdEvent = await executeGraphApiCall(
        () => axios.post<Event>(
          `https://graph.microsoft.com/v1.0/users/${microsoftUserId}/calendars/${calendarId}/events`,
          event,
          {
            headers: this.buildHeaders(accessToken),
          }
        ),
        {
          logger: this.logger,
          resourceName: `create event in calendar ${calendarId} for user ${microsoftUserId}`,
          maxRetries: 7,
        }
      );

      if (!createdEvent?.data) {
        throw new Error('Event creation returned no data');
      }

      return { event: createdEvent.data };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`[createEvent] Failed: ${errorMessage}`);
      throw new Error(`Failed to create calendar event: ${errorMessage}`);
    }
  }

  /**
   * Update an existing event in a user's calendar.
   *
   * @param eventId - The ID of the event to update
   * @param updates - Partial Event object with fields to update
   * @param tenantId - Microsoft tenant ID (Azure AD tenant GUID)
   * @param microsoftUserId - Microsoft Graph user ID
   * @param calendarId - Calendar ID where the event exists
   * @returns The updated event data
   */
  async updateEvent(
    eventId: string,
    updates: Partial<Event>,
    tenantId: string,
    microsoftUserId: string,
    calendarId: string,
  ): Promise<{ event: Event }> {
    try {
      this.logger.log(
        `[updateEvent] Updating event ${eventId} in calendar ${calendarId} for user ${microsoftUserId}`
      );

      // Get app-only access token for the tenant
      const accessToken = await this.appOnlyAuthService.getAccessToken(tenantId);

      const updatedEvent = await executeGraphApiCall(
        () => axios.patch<Event>(
          `https://graph.microsoft.com/v1.0/users/${microsoftUserId}/calendars/${calendarId}/events/${eventId}`,
          updates,
          {
            headers: this.buildHeaders(accessToken),
          }
        ),
        {
          logger: this.logger,
          resourceName: `update event ${eventId} for user ${microsoftUserId}`,
          maxRetries: 7,
        }
      );

      if (!updatedEvent?.data) {
        throw new Error('Event update returned no data');
      }

      return { event: updatedEvent.data };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`[updateEvent] Failed: ${errorMessage}`);
      throw new Error(`Failed to update calendar event: ${errorMessage}`);
    }
  }

  /**
   * Get a single event by its ID from a user's calendar.
   *
   * @param tenantId - Microsoft tenant ID (Azure AD tenant GUID)
   * @param microsoftUserId - Microsoft Graph user ID
   * @param eventId - Event ID to fetch
   * @returns Event object or null if not found (404)
   */
  async getEventById(
    tenantId: string,
    microsoftUserId: string,
    eventId: string,
  ): Promise<Event | null> {
    try {
      this.logger.debug(`[getEventById] Fetching event ${eventId} for user ${microsoftUserId}`);

      // Get app-only access token for the tenant
      const accessToken = await this.appOnlyAuthService.getAccessToken(tenantId);

      const event = await executeGraphApiCall(
        () => axios.get<Event>(
          `https://graph.microsoft.com/v1.0/users/${microsoftUserId}/events/${eventId}`,
          {
            headers: this.buildHeaders(accessToken),
          }
        ),
        {
          logger: this.logger,
          resourceName: `get event ${eventId} for user ${microsoftUserId}`,
          maxRetries: 7,
          return404AsNull: true,
        }
      );

      return event?.data ?? null;
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`[getEventById] Failed: ${errorMessage}`);
      throw new Error(`Failed to get calendar event: ${errorMessage}`);
    }
  }

  /**
   * Delete an event from a user's calendar.
   *
   * @param tenantId - Microsoft tenant ID (Azure AD tenant GUID)
   * @param microsoftUserId - Microsoft Graph user ID
   * @param eventId - The ID of the event to delete
   * @param calendarId - Calendar ID where the event exists
   */
  async deleteEvent(
    tenantId: string,
    microsoftUserId: string,
    eventId: string,
    calendarId: string,
  ): Promise<void> {
    try {
      this.logger.log(
        `[deleteEvent] Deleting event ${eventId} from calendar ${calendarId} for user ${microsoftUserId}`
      );

      // Get app-only access token for the tenant
      const accessToken = await this.appOnlyAuthService.getAccessToken(tenantId);

      await executeGraphApiCall(
        () => axios.delete(
          `https://graph.microsoft.com/v1.0/users/${microsoftUserId}/calendars/${calendarId}/events/${eventId}`,
          {
            headers: this.buildHeaders(accessToken),
          }
        ),
        {
          logger: this.logger,
          resourceName: `delete event ${eventId} for user ${microsoftUserId}`,
          maxRetries: 7,
          return404AsNull: true, // Treat already deleted as success
        }
      );
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`[deleteEvent] Failed: ${errorMessage}`);
      throw new Error(`Failed to delete calendar event: ${errorMessage}`);
    }
  }

  /**
   * Create multiple events in a single batch request.
   * Uses Microsoft Graph $batch API for efficient batch creation.
   *
   * @param events - Array of event objects to create
   * @param tenantId - Microsoft tenant ID (Azure AD tenant GUID)
   * @param microsoftUserId - Microsoft Graph user ID
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
    tenantId: string,
    microsoftUserId: string,
    calendarId: string,
  ): Promise<{ index: number; success: boolean; event?: Event; error?: string }[]> {
    if (events.length === 0) {
      return [];
    }

    const results: { index: number; success: boolean; event?: Event; error?: string }[] = [];
    const BATCH_SIZE = 20;

    try {
      // Get app-only access token for the tenant
      const accessToken = await this.appOnlyAuthService.getAccessToken(tenantId);

      for (let i = 0; i < events.length; i += BATCH_SIZE) {
        const batchEvents = events.slice(i, i + BATCH_SIZE);

        const batchPayload: BatchRequestPayload = {
          requests: batchEvents.map((event, index) => ({
            id: `${index}`,
            method: 'POST',
            url: `/users/${microsoftUserId}/calendars/${calendarId}/events`,
            body: event,
            headers: {
              'Content-Type': 'application/json',
              'Prefer': 'IdType="ImmutableId"',
            },
          })),
        };

        this.logger.log(
          `[createBatchEvents] Creating batch of ${batchEvents.length} events for user ${microsoftUserId}`
        );

        try {
          const response = await executeGraphApiCall(
            () => axios.post<BatchResponsePayload<Event>>(
              'https://graph.microsoft.com/v1.0/$batch',
              batchPayload,
              {
                headers: this.buildHeaders(accessToken),
              }
            ),
            {
              logger: this.logger,
              resourceName: `batch create ${batchEvents.length} events for user ${microsoftUserId}`,
              maxRetries: 7,
            }
          );

          if (!response?.data) {
            throw new Error('Batch request returned null response');
          }

          response.data.responses.forEach((batchResponse, batchIndex) => {
            const globalIndex = i + batchIndex;

            if (batchResponse.status === 201) {
              results.push({
                index: globalIndex,
                success: true,
                event: batchResponse.body,
              });
            } else {
              results.push({
                index: globalIndex,
                success: false,
                error: `HTTP ${batchResponse.status}: ${JSON.stringify(batchResponse.body)}`,
              });
            }
          });
        } catch (error) {
          const errorMessage = error instanceof Error ? error.message : 'Unknown error';
          batchEvents.forEach((_, batchIndex) => {
            results.push({
              index: i + batchIndex,
              success: false,
              error: `Batch request failed: ${errorMessage}`,
            });
          });
        }
      }

      const successCount = results.filter(r => r.success).length;
      this.logger.log(
        `[createBatchEvents] Completed: ${successCount}/${events.length} succeeded`
      );

      return results;
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`[createBatchEvents] Failed: ${errorMessage}`);
      return events.map((_, index) => ({
        index,
        success: false,
        error: `Batch creation failed: ${errorMessage}`,
      }));
    }
  }

  /**
   * Update multiple events in a single batch request.
   *
   * @param updates - Array of update objects with eventId and fields to update
   * @param tenantId - Microsoft tenant ID (Azure AD tenant GUID)
   * @param microsoftUserId - Microsoft Graph user ID
   * @param calendarId - Calendar ID
   * @returns Results array with success/failure for each update
   */
  async updateBatchEvents(
    updates: Array<{ eventId: string; updates: Partial<Event> }>,
    tenantId: string,
    microsoftUserId: string,
    calendarId: string,
  ): Promise<{ index: number; success: boolean; event?: Event; error?: string }[]> {
    if (updates.length === 0) {
      return [];
    }

    const results: { index: number; success: boolean; event?: Event; error?: string }[] = [];
    const BATCH_SIZE = 20;

    try {
      // Get app-only access token for the tenant
      const accessToken = await this.appOnlyAuthService.getAccessToken(tenantId);

      for (let i = 0; i < updates.length; i += BATCH_SIZE) {
        const batchUpdates = updates.slice(i, i + BATCH_SIZE);

        const batchPayload: BatchRequestPayload = {
          requests: batchUpdates.map((update, index) => ({
            id: `${index}`,
            method: 'PATCH',
            url: `/users/${microsoftUserId}/calendars/${calendarId}/events/${update.eventId}`,
            body: update.updates,
            headers: {
              'Content-Type': 'application/json',
              'Prefer': 'IdType="ImmutableId"',
            },
          })),
        };

        this.logger.log(
          `[updateBatchEvents] Updating batch of ${batchUpdates.length} events for user ${microsoftUserId}`
        );

        try {
          const response = await executeGraphApiCall(
            () => axios.post<BatchResponsePayload<Event>>(
              'https://graph.microsoft.com/v1.0/$batch',
              batchPayload,
              {
                headers: this.buildHeaders(accessToken),
              }
            ),
            {
              logger: this.logger,
              resourceName: `batch update ${batchUpdates.length} events for user ${microsoftUserId}`,
              maxRetries: 7,
            }
          );

          if (!response?.data) {
            throw new Error('Batch request returned null response');
          }

          response.data.responses.forEach((batchResponse, batchIndex) => {
            const globalIndex = i + batchIndex;

            if (batchResponse.status === 200) {
              results.push({
                index: globalIndex,
                success: true,
                event: batchResponse.body,
              });
            } else {
              results.push({
                index: globalIndex,
                success: false,
                error: `HTTP ${batchResponse.status}: ${JSON.stringify(batchResponse.body)}`,
              });
            }
          });
        } catch (error) {
          const errorMessage = error instanceof Error ? error.message : 'Unknown error';
          batchUpdates.forEach((_, batchIndex) => {
            results.push({
              index: i + batchIndex,
              success: false,
              error: `Batch request failed: ${errorMessage}`,
            });
          });
        }
      }

      const successCount = results.filter(r => r.success).length;
      this.logger.log(
        `[updateBatchEvents] Completed: ${successCount}/${updates.length} succeeded`
      );

      return results;
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`[updateBatchEvents] Failed: ${errorMessage}`);
      return updates.map((_, index) => ({
        index,
        success: false,
        error: `Batch update failed: ${errorMessage}`,
      }));
    }
  }

  /**
   * Delete multiple events in a single batch request.
   *
   * @param eventIds - Array of event IDs to delete
   * @param tenantId - Microsoft tenant ID (Azure AD tenant GUID)
   * @param microsoftUserId - Microsoft Graph user ID
   * @param calendarId - Calendar ID
   * @returns Results array with success/failure for each event
   */
  async deleteBatchEvents(
    eventIds: string[],
    tenantId: string,
    microsoftUserId: string,
    calendarId: string,
  ): Promise<{ id: string; success: boolean; error?: string }[]> {
    if (eventIds.length === 0) {
      return [];
    }

    const results: { id: string; success: boolean; error?: string }[] = [];
    const BATCH_SIZE = 20;

    try {
      // Get app-only access token for the tenant
      const accessToken = await this.appOnlyAuthService.getAccessToken(tenantId);

      for (let i = 0; i < eventIds.length; i += BATCH_SIZE) {
        const batchEventIds = eventIds.slice(i, i + BATCH_SIZE);

        const batchPayload: BatchRequestPayload = {
          requests: batchEventIds.map((eventId, index) => ({
            id: `${index}`,
            method: 'DELETE',
            url: `/users/${microsoftUserId}/calendars/${calendarId}/events/${eventId}`,
            headers: {
              'Prefer': 'IdType="ImmutableId"',
            },
          })),
        };

        this.logger.log(
          `[deleteBatchEvents] Deleting batch of ${batchEventIds.length} events for user ${microsoftUserId}`
        );

        try {
          const response = await executeGraphApiCall(
            () => axios.post<BatchResponsePayload>(
              'https://graph.microsoft.com/v1.0/$batch',
              batchPayload,
              {
                headers: this.buildHeaders(accessToken),
              }
            ),
            {
              logger: this.logger,
              resourceName: `batch delete ${batchEventIds.length} events for user ${microsoftUserId}`,
              maxRetries: 7,
            }
          );

          if (!response?.data) {
            throw new Error('Batch request returned null response');
          }

          response.data.responses.forEach((batchResponse, index) => {
            const eventId = batchEventIds[index];

            // 204 (No Content) or 404 (already deleted) = success
            if (batchResponse.status === 204 || batchResponse.status === 404) {
              results.push({ id: eventId, success: true });
            } else {
              results.push({
                id: eventId,
                success: false,
                error: `HTTP ${batchResponse.status}: ${JSON.stringify(batchResponse.body)}`,
              });
            }
          });
        } catch (error) {
          const errorMessage = error instanceof Error ? error.message : 'Unknown error';
          batchEventIds.forEach(eventId => {
            results.push({
              id: eventId,
              success: false,
              error: `Batch request failed: ${errorMessage}`,
            });
          });
        }
      }

      const successCount = results.filter(r => r.success).length;
      this.logger.log(
        `[deleteBatchEvents] Completed: ${successCount}/${eventIds.length} succeeded`
      );

      return results;
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`[deleteBatchEvents] Failed: ${errorMessage}`);
      return eventIds.map(id => ({
        id,
        success: false,
        error: `Batch deletion failed: ${errorMessage}`,
      }));
    }
  }

  /**
   * Stream calendar events for a user using calendarView.
   *
   * This method uses an async generator pattern to stream events in batches,
   * minimizing memory usage for large calendars.
   *
   * @param tenantId - Microsoft tenant ID (Azure AD tenant GUID)
   * @param microsoftUserId - Microsoft Graph user ID
   * @param options - Optional configuration
   * @yields Chunks of events
   */
  async *streamEvents(
    tenantId: string,
    microsoftUserId: string,
    options?: {
      startDate?: Date;
      endDate?: Date;
      batchSize?: number;
    },
  ): AsyncGenerator<Event[], void, unknown> {
    const batchSize = options?.batchSize ?? 100;
    const now = new Date();
    const dateInterval = 5 * 365 * 24 * 60 * 60 * 1000; // 5 years

    const startDate = options?.startDate ?? now;
    const endDate = options?.endDate ?? new Date(Date.now() + dateInterval);

    const startDateStr = startDate.toISOString();
    const endDateStr = endDate.toISOString();

    this.logger.log(
      `[streamEvents] Starting event stream for user ${microsoftUserId} from ${startDateStr} to ${endDateStr}`
    );

    try {
      // Get app-only access token for the tenant
      const accessToken = await this.appOnlyAuthService.getAccessToken(tenantId);

      let nextLink: string | undefined =
        `https://graph.microsoft.com/v1.0/users/${microsoftUserId}/calendarView?startDateTime=${startDateStr}&endDateTime=${endDateStr}&$orderby=start/dateTime&$top=${batchSize}`;

      const buffer: Event[] = [];
      let totalFetched = 0;

      while (nextLink) {
        const response = await executeGraphApiCall(
          () => axios.get<{ value: Event[]; '@odata.nextLink'?: string }>(
            nextLink as string,
            {
              headers: this.buildHeaders(accessToken),
            }
          ),
          {
            logger: this.logger,
            resourceName: `calendarView for user ${microsoftUserId}`,
          }
        );

        if (!response?.data.value) {
          break;
        }

        const items = response.data.value;
        buffer.push(...items);
        totalFetched += items.length;

        // Yield when buffer reaches batch size
        while (buffer.length >= batchSize) {
          const chunk = buffer.splice(0, batchSize);
          yield chunk;
        }

        nextLink = response.data['@odata.nextLink'];
      }

      // Yield remaining items
      if (buffer.length > 0) {
        yield buffer;
      }

      this.logger.log(
        `[streamEvents] Completed streaming ${totalFetched} events for user ${microsoftUserId}`
      );

      // Emit completion event
      this.eventEmitter.emit(OutlookEventTypes.IMPORT_COMPLETED, {
        tenantId,
        microsoftUserId,
        totalEvents: totalFetched,
        isTenantWide: true,
      });
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(
        `[streamEvents] Error streaming events for user ${microsoftUserId}: ${errorMessage}`
      );
      throw error;
    }
  }

  /**
   * Get multiple events by IDs in a single batch request.
   *
   * @param eventIds - Array of event IDs to fetch (max 20)
   * @param tenantId - Microsoft tenant ID (Azure AD tenant GUID)
   * @param microsoftUserId - Microsoft Graph user ID
   * @returns Array of successfully fetched events
   */
  async getEventsBatch(
    eventIds: string[],
    tenantId: string,
    microsoftUserId: string,
  ): Promise<Event[]> {
    if (eventIds.length === 0) {
      return [];
    }

    if (eventIds.length > 20) {
      this.logger.warn(
        `[getEventsBatch] Called with ${eventIds.length} events, exceeding limit. Only first 20 will be fetched.`
      );
    }

    try {
      // Get app-only access token for the tenant
      const accessToken = await this.appOnlyAuthService.getAccessToken(tenantId);

      const batchPayload: BatchRequestPayload = {
        requests: eventIds.slice(0, 20).map((eventId, index) => ({
          id: `${index}`,
          method: 'GET',
          url: `/users/${microsoftUserId}/events/${eventId}`,
          headers: {
            'Prefer': 'IdType="ImmutableId"',
          },
        })),
      };

      this.logger.debug(
        `[getEventsBatch] Fetching ${batchPayload.requests.length} events for user ${microsoftUserId}`
      );

      const response = await executeGraphApiCall(
        () => axios.post<BatchResponsePayload<Event>>(
          'https://graph.microsoft.com/v1.0/$batch',
          batchPayload,
          {
            headers: this.buildHeaders(accessToken),
          }
        ),
        {
          logger: this.logger,
          resourceName: `batch get events for user ${microsoftUserId}`,
          maxRetries: 7,
        }
      );

      if (!response?.data) {
        return [];
      }

      const events: Event[] = [];

      for (const batchResponse of response.data.responses) {
        if (batchResponse.status === 200) {
          events.push(batchResponse.body);
        } else if (batchResponse.status === 404) {
          this.logger.debug(
            `[getEventsBatch] Event not found (404), likely deleted`
          );
        } else {
          this.logger.warn(
            `[getEventsBatch] Event fetch failed with status ${batchResponse.status}`
          );
        }
      }

      return events;
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`[getEventsBatch] Failed: ${errorMessage}`);
      throw error;
    }
  }

  /**
   * Clear the calendar ID cache for a specific user or all users.
   *
   * @param microsoftUserId - Optional user ID to clear cache for (clears all if not specified)
   */
  clearCache(microsoftUserId?: string): void {
    if (microsoftUserId) {
      this.logger.log(`[clearCache] Clearing calendar cache for user ${microsoftUserId}`);
      // Note: Cache key includes tenantId, so we clear entire cache for simplicity
      this.defaultCalendarIdCache.clear();
    } else {
      this.logger.log('[clearCache] Clearing entire calendar cache');
      this.defaultCalendarIdCache.clear();
    }
  }
}
