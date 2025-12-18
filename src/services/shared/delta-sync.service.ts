import { Injectable, Logger } from "@nestjs/common";
import { Client } from "@microsoft/microsoft-graph-client";
import { OutlookDeltaLinkRepository } from "../../repositories/outlook-delta-link.repository";
import { ResourceType } from "../../enums/resource-type.enum";
import { Event, Message } from "../../types";
import { delay, retryWithBackoff } from "../../utils/retry.util";

export interface DeltaItem {
  lastModifiedDateTime?: string;
  createdDateTime?: string;
  id?: string;
  "@removed"?: {
    reason: "changed" | "deleted";
  };
}

export type DeltaEvent = Event & DeltaItem;
export type DeltaMessage = Message & DeltaItem;

export interface DeltaResponse<T> {
  "@odata.nextLink"?: string;
  "@odata.deltaLink"?: string;
  value: T[];
}

export class DeltaSyncError extends Error {
  constructor(
    message: string,
    public readonly code: string,
    public readonly statusCode: number
  ) {
    super(message);
    this.name = "DeltaSyncError";
  }
}

@Injectable()
export class DeltaSyncService {
  private readonly logger = new Logger(DeltaSyncService.name);
  private readonly MAX_RETRIES = 3;
  private readonly RETRY_DELAY_MS = 1000; // 1 second

  constructor(
    private readonly deltaLinkRepository: OutlookDeltaLinkRepository
  ) {}

  private handleDeltaResponse<T extends DeltaItem>(
    response: DeltaResponse<T>,
    userId: number,
    resourceType: ResourceType
  ): void {
    // Handle sync reset (410 Gone)
    if (response["@odata.deltaLink"]?.includes("$deltatoken=")) {
      this.logger.log(
        `Sync reset detected for user ${userId}, resource ${resourceType} with ${response.value.length} changes`
      );
      // Note: Delta link will be saved after processing all changes in fetchAndSortChanges
      // Saving it here would skip the changes in the current response
    }

    // Handle token expiration
    if (response["@odata.deltaLink"]) {
      const tokenExpiry = this.calculateTokenExpiry(resourceType);
      this.logger.log(
        `Delta token will expire at ${tokenExpiry.toISOString()}`
      );
    }
  }

  private calculateTokenExpiry(resourceType: ResourceType): Date {
    const now = new Date();
    // Directory objects and education objects have 7-day expiry
    if (resourceType === ResourceType.CALENDAR) {
      // For Outlook entities, we'll use a conservative 6-day expiry
      // since the actual limit depends on internal cache size
      return new Date(now.getTime() + 6 * 24 * 60 * 60 * 1000);
    }
    // Default to 7 days for other resources
    return new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);
  }

  private handleReplays<T extends DeltaItem>(items: T[]): T[] {
    // Use a Map to deduplicate items by ID
    const uniqueItems = new Map<string, T>();

    for (const item of items) {
      if (item.id) {
        // If item exists and has @removed, keep the removal
        if (item["@removed"]) {
          uniqueItems.set(item.id, item);
        }
        // If item exists and is not removed, update it
        else if (
          !uniqueItems.has(item.id) ||
          !uniqueItems.get(item.id)?.["@removed"]
        ) {
          uniqueItems.set(item.id, item);
        }
      }
    }

    return Array.from(uniqueItems.values());
  }

  /**
   * Sorts delta items by their modification/creation timestamp
   * @param items Items to sort
   * @returns Sorted items (oldest to newest)
   */
  private sortDeltaItems(items: DeltaItem[]): DeltaItem[] {
    return items.sort((a, b) => {
      const aTime = a.lastModifiedDateTime ?? a.createdDateTime ?? "";
      const bTime = b.lastModifiedDateTime ?? b.createdDateTime ?? "";
      return new Date(aTime).getTime() - new Date(bTime).getTime();
    });
  }

  /**
   * Fetches all pages from a delta query and returns items with the final delta link
   * This is the core pagination logic shared by both initialization and incremental sync
   * 
   * features:
   * - pagination
   * - retry with exponential backoff
   * - deduplication
   * - logging
   * - error handling
   * - rate limiting
   *
   * @param client Microsoft Graph client
   * @param startUrl Initial URL to start fetching from
   * @param userId User ID for logging
   * @returns Object with all items and the final delta link
   */
  private async fetchAllDeltaPages(
    client: Client,
    startUrl: string,
    userId: number
  ): Promise<{ items: DeltaItem[]; deltaLink: string | null }> {
    let response: DeltaResponse<DeltaItem> = {
      "@odata.nextLink": startUrl,
      value: [],
    };

    let lastDeltaLink: string | null = null;
    const allItems: DeltaItem[] = [];
    let pageCount = 0;

    // Fetch all pages until we get the delta link
    while (response["@odata.nextLink"]) {
      const nextLink = response["@odata.nextLink"];
      pageCount++;

      this.logger.debug(`[fetchAllDeltaPages] Fetching page ${pageCount} for user ${userId}`);

      response = (await retryWithBackoff(
        () => client.api(nextLink).get(),
        {
          maxRetries: this.MAX_RETRIES,
          retryDelayMs: this.RETRY_DELAY_MS,
        }
      )) as DeltaResponse<DeltaItem>;

      this.logger.debug(`[fetchAllDeltaPages] Received ${response.value.length} items in page ${pageCount}`);

      // Check if we got the delta link (indicates we've reached the end)
      if (response["@odata.deltaLink"]) {
        lastDeltaLink = this.getDeltaLink(response);
        this.logger.log(`[fetchAllDeltaPages] Reached end after ${pageCount} pages, got delta link`);
      }

      // Fetch full event details for each item (skip deleted items)
      const eventDetails = await Promise.all(
        response.value.map((item) =>
          item["@removed"]
            ? Promise.resolve(item)
            : (client.api(`/me/events/${item.id}`).get() as Promise<DeltaItem>)
        )
      );
      allItems.push(...eventDetails);

      await delay(200); // Slight delay to avoid hitting rate limits
    }

    this.logger.log(`[fetchAllDeltaPages] Fetched total of ${allItems.length} items across ${pageCount} pages for user ${userId}`);

    return { items: allItems, deltaLink: lastDeltaLink };
  }

  /**
   * Fetches and sorts delta changes for any resource type
   * @param client Microsoft Graph client
   * @param requestUrl Initial request URL
   * @param userId User ID
   * @param forceReset Force reset of delta link
   * @param dateRange Optional date range for calendar delta queries (only used on initialization)
   * @returns Object containing sorted items and delta link for next sync
   */
  async fetchAndSortChanges(
    client: Client,
    requestUrl: string,
    userId: string,
    forceReset: boolean = false,
    dateRange?: {
      startDate: Date;
      endDate: Date;
    }
  ): Promise<DeltaItem[]> {
    let startLink = await this.deltaLinkRepository.getDeltaLink(
      Number(userId),
      ResourceType.CALENDAR
    );

    this.logger.log(`[fetchAndSortChanges] startLink: ${startLink} forceReset: ${forceReset} dateRange: ${JSON.stringify(dateRange)}`);

    // Force reset if requested (e.g., on reconnection)
    if (forceReset && startLink) {
      this.logger.log(`[fetchAndSortChanges] Force reset requested, deleting existing delta link for user ${userId}`);
      await this.deltaLinkRepository.deleteDeltaLink(Number(userId), ResourceType.CALENDAR);
      startLink = null;
    }

    // If no delta link exists, initialize from "now" and return current items
    // This fetches all current events and establishes the delta link baseline
    if (!startLink) {
      this.logger.log(`[fetchAndSortChanges] No delta link found for user ${userId}, initializing from current point`);
      const result = await this.initializeDeltaLink(
        client,
        requestUrl,
        Number(userId),
        ResourceType.CALENDAR,
        dateRange
      );

      // Sort items before returning
      return this.sortDeltaItems(result);
    }

    // Incremental sync: fetch changes since last delta link
    this.logger.debug(`[fetchAndSortChanges] Starting incremental sync with existing delta link for user ${userId}`);

    const { items: allItems, deltaLink: lastDeltaLink } = await this.fetchAllDeltaPages(
      client,
      startLink,
      Number(userId)
    );

    // Save the delta link for incremental syncs (initialization already saves it)
    if (lastDeltaLink) {
      await this.saveDeltaLink(Number(userId), ResourceType.CALENDAR, lastDeltaLink);
      this.logger.log(`[fetchAndSortChanges] Saved delta link after fetching ${allItems.length} changes for user ${userId}`);
    }

    // Sort and return items
    return this.sortDeltaItems(allItems);
  }

  /**
   * Initialize delta link from current point in time, returning all current items.
   * This establishes a baseline and returns items so they can be processed.
   *
   * @param client Microsoft Graph client
   * @param requestUrl Initial delta request URL (e.g., "/me/events/delta")
   * @param userId User ID
   * @param resourceType Resource type (e.g., CALENDAR)
   * @param dateRange Optional date range for calendar delta queries
   * @param dateRange.startDate Start date for sync window
   * @param dateRange.endDate End date for sync window
   * @returns Object with items and delta link
   */
  async initializeDeltaLink(
    client: Client,
    requestUrl: string,
    userId: number,
    resourceType: ResourceType,
    dateRange?: {
      startDate: Date;
      endDate: Date;
    }
  ): Promise<DeltaItem[]> {
    this.logger.log(`[initializeDeltaLink] Initializing delta link and fetching current items for user ${userId}`);

    let urlWithDateRange = requestUrl;

    // Add date range parameters if provided
    if (dateRange) {
      const { startDate, endDate } = dateRange;
      urlWithDateRange = `${requestUrl}?startDateTime=${startDate.toISOString()}&endDateTime=${endDate.toISOString()}`;
      this.logger.log(`[initializeDeltaLink] Using date range: ${startDate.toISOString()} to ${endDate.toISOString()}`);
    }

    // Fetch all delta pages using shared function
    const { items: allItems, deltaLink: lastDeltaLink } = await this.fetchAllDeltaPages(
      client,
      urlWithDateRange,
      userId
    );

    if (!lastDeltaLink) {
      throw new Error('Failed to initialize delta link - no delta link received from Microsoft Graph');
    }

    // Save the delta link for future syncs
    await this.saveDeltaLink(userId, resourceType, lastDeltaLink);
    this.logger.log(`[initializeDeltaLink] Delta link initialized and saved for user ${userId}, returning ${allItems.length} items to process`);

    return allItems;
  }

  /**
   * Save delta link for next sync
   * Should be called AFTER all events from fetchAndSortChanges have been processed
   */
  async saveDeltaLink(
    userId: number,
    resourceType: ResourceType,
    deltaLink: string
  ): Promise<void> {
    await this.deltaLinkRepository.saveDeltaLink(userId, resourceType, deltaLink);
    this.logger.debug(`[saveDeltaLink] Saved delta link for user ${userId}, resource ${resourceType}`);
  }

  /**
   * Gets the delta link from the response
   * @param response Delta response
   * @returns Delta link or null
   */
  getDeltaLink<T>(response: DeltaResponse<T>): string | null {
    return response["@odata.deltaLink"] || null;
  }
}
