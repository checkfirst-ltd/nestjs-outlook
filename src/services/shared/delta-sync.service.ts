import { Injectable, Logger } from "@nestjs/common";
import { Client } from "@microsoft/microsoft-graph-client";
import { OutlookDeltaLinkRepository } from "../../repositories/outlook-delta-link.repository";
import { ResourceType } from "../../enums/resource-type.enum";
import { Event, Message } from "../../types";

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

  private async delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  private async retryWithBackoff<T>(
    operation: () => Promise<T>,
    retryCount = 0
  ): Promise<T> {
    try {
      return await operation();
    } catch (error) {
      if (retryCount >= this.MAX_RETRIES) {
        throw error;
      }

      // Calculate exponential backoff delay
      const delayMs = this.RETRY_DELAY_MS * Math.pow(2, retryCount);
      await this.delay(delayMs);
      return this.retryWithBackoff(operation, retryCount + 1);
    }
  }

  private handleDeltaResponse<T extends DeltaItem>(
    response: DeltaResponse<T>,
    userId: number,
    resourceType: ResourceType
  ): void {
    // Handle sync reset (410 Gone)
    if (response["@odata.deltaLink"]?.includes("$deltatoken=")) {
      this.logger.log(
        `Sync reset detected for user ${userId}, resource ${resourceType}`
      );
      // Clear the delta link to force a full sync
      void this.deltaLinkRepository.saveDeltaLink(
        userId,
        resourceType,
        this.getDeltaLink(response) ?? ""
      );
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
   * Fetches and sorts delta changes for any resource type
   * @param client Microsoft Graph client
   * @param requestUrl Initial request URL
   * @returns Array of items sorted by lastModifiedDateTime
   */
  async fetchAndSortChanges<T extends DeltaItem>(
    client: Client,
    requestUrl: string,
    userId: string
  ): Promise<T[]> {
    const allItems: T[] = [];
    let response: DeltaResponse<T> = {
      "@odata.nextLink":
        (await this.deltaLinkRepository.getDeltaLink(
          Number(userId),
          ResourceType.CALENDAR
        )) ?? requestUrl,
      value: [],
    };

    // Fetch all pages of changes
    while (response["@odata.nextLink"]) {
      const nextLink = response["@odata.nextLink"];
      response = (await client.api(nextLink).get()) as DeltaResponse<T>;
      this.handleDeltaResponse(response, Number(userId), ResourceType.CALENDAR);
      const eventDetails = await Promise.all(
        response.value.map((item) =>
          item["@removed"]
            ? Promise.resolve(item)
            : (client.api(`/me/events/${item.id}`).get() as Promise<T>)
        )
      );
      allItems.push(...eventDetails);
      await this.delay(200); // Slight delay to avoid hitting rate limits
    }

    // Sort by lastModifiedDateTime (fallback to createdDateTime)
    return allItems.sort((a, b) => {
      const aTime = a.lastModifiedDateTime ?? a.createdDateTime ?? "";
      const bTime = b.lastModifiedDateTime ?? b.createdDateTime ?? "";
      return new Date(aTime).getTime() - new Date(bTime).getTime();
    });
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
