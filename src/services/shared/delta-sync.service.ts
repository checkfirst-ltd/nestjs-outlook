import { Injectable, Logger } from "@nestjs/common";
import { Client } from "@microsoft/microsoft-graph-client";
import { OutlookDeltaLinkRepository } from "../../repositories/outlook-delta-link.repository";
import { ResourceType } from "../../enums/resource-type.enum";
import { Event, Message } from "../../types";
import { delay, retryWithBackoff } from "../../utils/retry.util";
import { UserIdConverterService } from "./user-id-converter.service";

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
    private readonly deltaLinkRepository: OutlookDeltaLinkRepository,
    private readonly userIdConverter: UserIdConverterService
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
   * Core async generator that fetches delta pages one at a time
   * This is the foundational logic used by both batch and streaming approaches
   *
   * Features:
   * - Pagination through @odata.nextLink
   * - Retry with exponential backoff
   * - Fetch full event details
   * - Rate limiting between pages
   * - Logging
   *
   * @param client Microsoft Graph client
   * @param startUrl Initial URL to start fetching from
   * @param userId User ID for logging
   * @yields Object with page items, delta link (if last page), and isLastPage flag
   */
  private async *fetchDeltaPagesCore(
    client: Client,
    startUrl: string,
    userId: number
  ): AsyncGenerator<
    { items: DeltaItem[]; deltaLink: string | null; isLastPage: boolean },
    void,
    unknown
  > {
    let currentUrl: string = startUrl;
    let pageCount = 0;

    while (currentUrl) {
      pageCount++;
      this.logger.debug(`[fetchDeltaPagesCore] Fetching page ${pageCount} for user ${userId}`);

      // Fetch page with retry logic
      const response = (await retryWithBackoff(
        () => client.api(currentUrl).get(),
        {
          maxRetries: this.MAX_RETRIES,
          retryDelayMs: this.RETRY_DELAY_MS,
        }
      )) as DeltaResponse<DeltaItem>;

      this.logger.debug(`[fetchDeltaPagesCore] Received ${response.value.length} items in page ${pageCount}`);

      // Fetch full event details for each item (skip deleted items)
      const eventDetails = await Promise.all(
        response.value.map((item) =>
          item["@removed"]
            ? Promise.resolve(item)
            : (client.api(`/me/events/${item.id}`).get() as Promise<DeltaItem>)
        )
      );

      // Check if we got the delta link (indicates last page)
      const deltaLink = response["@odata.deltaLink"]
        ? this.getDeltaLink(response)
        : null;
      const isLastPage = deltaLink !== null;

      if (isLastPage) {
        this.logger.log(`[fetchDeltaPagesCore] Reached last page (${pageCount}) with delta link for user ${userId}`);
      }

      // Yield this page
      yield {
        items: eventDetails,
        deltaLink,
        isLastPage,
      };

      // Update URL for next iteration
      currentUrl = response["@odata.nextLink"] || "";

      // Rate limiting
      if (currentUrl) {
        await delay(200);
      }
    }
  }

  /**
   * Fetches all pages from a delta query and returns items with the final delta link
   * This method collects all items into memory before returning
   * For streaming alternative, use streamDeltaChanges()
   *
   * Features:
   * - Pagination
   * - Retry with exponential backoff
   * - Fetch full event details
   * - Rate limiting
   * - Logging
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
    const allItems: DeltaItem[] = [];
    let lastDeltaLink: string | null = null;
    let pageCount = 0;

    // Consume core generator and collect all items
    for await (const page of this.fetchDeltaPagesCore(client, startUrl, userId)) {
      allItems.push(...page.items);
      pageCount++;

      if (page.isLastPage && page.deltaLink) {
        lastDeltaLink = page.deltaLink;
      }
    }

    this.logger.log(`[fetchAllDeltaPages] Collected ${allItems.length} items across ${pageCount} pages for user ${userId}`);

    return { items: allItems, deltaLink: lastDeltaLink };
  }

  /**
   * Fetches and sorts delta changes for any resource type
   * @param client Microsoft Graph client
   * @param requestUrl Initial request URL
   * @param externalUserId External user ID from the host application
   * @param forceReset Force reset of delta link
   * @param dateRange Optional date range for calendar delta queries (only used on initialization)
   * @returns Object containing sorted items and delta link for next sync
   */
  async fetchAndSortChanges(
    client: Client,
    requestUrl: string,
    externalUserId: string,
    forceReset: boolean = false,
    dateRange?: {
      startDate: Date;
      endDate: Date;
    }
  ): Promise<DeltaItem[]> {
    // Convert external ID to internal ID for database operations
    const internalUserId = await this.userIdConverter.externalToInternal(externalUserId);

    let startLink = await this.deltaLinkRepository.getDeltaLink(
      internalUserId,
      ResourceType.CALENDAR
    );

    this.logger.log(`[fetchAndSortChanges] startLink: ${startLink} forceReset: ${forceReset} dateRange: ${JSON.stringify(dateRange)}`);

    // Force reset if requested (e.g., on reconnection)
    if (forceReset && startLink) {
      this.logger.log(`[fetchAndSortChanges] Force reset requested, deleting existing delta link for user ${externalUserId}`);
      await this.deltaLinkRepository.deleteDeltaLink(internalUserId, ResourceType.CALENDAR);
      startLink = null;
    }

    // If no delta link exists, initialize from "now" and return current items
    // This fetches all current events and establishes the delta link baseline
    if (!startLink) {
      this.logger.log(`[fetchAndSortChanges] No delta link found for user ${externalUserId}, initializing from current point`);
      const result = await this.initializeDeltaLink(
        client,
        requestUrl,
        internalUserId,
        ResourceType.CALENDAR,
        dateRange
      );

      // Sort items before returning
      return this.sortDeltaItems(result);
    }

    // Incremental sync: fetch changes since last delta link
    this.logger.debug(`[fetchAndSortChanges] Starting incremental sync with existing delta link for user ${externalUserId}`);

    const { items: allItems, deltaLink: lastDeltaLink } = await this.fetchAllDeltaPages(
      client,
      startLink,
      internalUserId
    );

    // Save the delta link for incremental syncs (initialization already saves it)
    if (lastDeltaLink) {
      await this.saveDeltaLink(internalUserId, ResourceType.CALENDAR, lastDeltaLink);
      this.logger.log(`[fetchAndSortChanges] Saved delta link after fetching ${allItems.length} changes for user ${externalUserId}`);
    }

    // Sort and return items
    return this.sortDeltaItems(allItems);
  }

  /**
   * Streams delta changes using async generator (memory-efficient alternative to fetchAndSortChanges)
   * Yields sorted batches of items as each page is fetched from Microsoft Graph
   *
   * Benefits over fetchAndSortChanges:
   * - Memory efficient: O(page_size) instead of O(total_items)
   * - Faster time-to-first-item: Start processing after first page
   * - Better for large syncs: Handle 1000s of changes without loading all into memory
   *
   * @param client Microsoft Graph client
   * @param requestUrl Initial request URL
   * @param externalUserId External user ID from the host application
   * @param forceReset Force reset of delta link
   * @param dateRange Optional date range for calendar delta queries (only used on initialization)
   * @param saveDeltaLink Whether to save the delta link to database (default: true). Set to false for one-time windowed imports.
   * @yields Sorted batches of delta items (one batch per Microsoft Graph page)
   * @returns Final delta link (saved to database only if saveDeltaLink=true)
   */
  async *streamDeltaChanges(
    client: Client,
    requestUrl: string,
    externalUserId: string,
    forceReset: boolean = false,
    dateRange?: {
      startDate: Date;
      endDate: Date;
    },
    saveDeltaLink: boolean = true
  ): AsyncGenerator<DeltaItem[], string | null, unknown> {
    // Convert external ID to internal ID for database operations
    const internalUserId = await this.userIdConverter.externalToInternal(externalUserId);

    let startLink = await this.deltaLinkRepository.getDeltaLink(
      internalUserId,
      ResourceType.CALENDAR
    );

    this.logger.log(`[streamDeltaChanges] Starting stream for user ${externalUserId}, startLink: ${startLink ? 'exists' : 'none'}, forceReset: ${forceReset}`);

    // Force reset if requested (e.g., on reconnection)
    if (forceReset && startLink) {
      this.logger.log(`[streamDeltaChanges] Force reset requested, deleting existing delta link for user ${externalUserId}`);
      await this.deltaLinkRepository.deleteDeltaLink(internalUserId, ResourceType.CALENDAR);
      startLink = null;
    }

    // Determine the starting URL
    let urlToUse: string;
    let finalDeltaLink: string | null = null;

    if (!startLink) {
      // No delta link exists - initialize from "now"
      this.logger.log(`[streamDeltaChanges] No delta link found, initializing from current point for user ${externalUserId}`);

      // Build URL with date range if provided
      if (dateRange) {
        const { startDate, endDate } = dateRange;
        urlToUse = `${requestUrl}?startDateTime=${startDate.toISOString()}&endDateTime=${endDate.toISOString()}`;
        this.logger.log(`[streamDeltaChanges] Using date range: ${startDate.toISOString()} to ${endDate.toISOString()}`);
      } else {
        urlToUse = requestUrl;
      }
    } else {
      // Delta link exists - incremental sync
      this.logger.log(`[streamDeltaChanges] Using existing delta link for incremental sync for user ${externalUserId}`);
      urlToUse = startLink;
    }

    // Stream pages using core generator
    let pageCount = 0;
    for await (const page of this.fetchDeltaPagesCore(client, urlToUse, internalUserId)) {
      pageCount++;

      // Sort and yield this batch immediately
      const sortedBatch = this.sortDeltaItems(page.items);
      this.logger.log(`[streamDeltaChanges] Yielding page ${pageCount} with ${sortedBatch.length} sorted items for user ${externalUserId}`);

      yield sortedBatch;

      // Capture final delta link
      if (page.isLastPage && page.deltaLink) {
        finalDeltaLink = page.deltaLink;
      }
    }

    // Save delta link after streaming all pages (if requested)
    if (finalDeltaLink && saveDeltaLink) {
      await this.saveDeltaLink(internalUserId, ResourceType.CALENDAR, finalDeltaLink);
      this.logger.log(`[streamDeltaChanges] Saved delta link after streaming ${pageCount} pages for user ${externalUserId}`);
    } else if (finalDeltaLink && !saveDeltaLink) {
      this.logger.log(`[streamDeltaChanges] Delta link discarded (saveDeltaLink=false) after streaming ${pageCount} pages for user ${externalUserId}`);
    } else {
      this.logger.warn(`[streamDeltaChanges] No delta link received after streaming ${pageCount} pages for user ${externalUserId}`);
    }

    return finalDeltaLink;
  }

  /**
   * Initialize delta link from current point in time, returning all current items.
   * This establishes a baseline and returns items so they can be processed.
   *
   * @param client Microsoft Graph client
   * @param requestUrl Initial delta request URL (e.g., "/me/events/delta")
   * @param internalUserId Internal user ID (MicrosoftUser.id)
   * @param resourceType Resource type (e.g., CALENDAR)
   * @param dateRange Optional date range for calendar delta queries
   * @param dateRange.startDate Start date for sync window
   * @param dateRange.endDate End date for sync window
   * @returns Object with items and delta link
   * @private This method is for internal use only
   */
  async initializeDeltaLink(
    client: Client,
    requestUrl: string,
    internalUserId: number,
    resourceType: ResourceType,
    dateRange?: {
      startDate: Date;
      endDate: Date;
    }
  ): Promise<DeltaItem[]> {
    this.logger.log(`[initializeDeltaLink] Initializing delta link and fetching current items for user ${internalUserId}`);

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
      internalUserId
    );

    if (!lastDeltaLink) {
      throw new Error('Failed to initialize delta link - no delta link received from Microsoft Graph');
    }

    // Save the delta link for future syncs
    await this.saveDeltaLink(internalUserId, resourceType, lastDeltaLink);
    this.logger.log(`[initializeDeltaLink] Delta link initialized and saved for user ${internalUserId}, returning ${allItems.length} items to process`);

    return allItems;
  }

  /**
   * Save delta link for next sync
   * Should be called AFTER all events from fetchAndSortChanges have been processed
   * @param internalUserId - Internal user ID (MicrosoftUser.id)
   * @param resourceType - Resource type (e.g., CALENDAR)
   * @param deltaLink - The delta link from Microsoft Graph
   * @private This method is for internal use only
   */
  private async saveDeltaLink(
    internalUserId: number,
    resourceType: ResourceType,
    deltaLink: string
  ): Promise<void> {
    await this.deltaLinkRepository.saveDeltaLink(internalUserId, resourceType, deltaLink);
    this.logger.debug(`[saveDeltaLink] Saved delta link for user ${internalUserId}, resource ${resourceType}`);
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
