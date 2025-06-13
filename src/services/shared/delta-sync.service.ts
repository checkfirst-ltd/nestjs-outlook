import { Injectable, Logger } from '@nestjs/common';
import { Client } from '@microsoft/microsoft-graph-client';
import { OutlookDeltaLinkRepository } from '../../repositories/outlook-delta-link.repository';
import { ResourceType } from '../../enums/resource-type.enum';

export interface DeltaItem {
  lastModifiedDateTime?: string;
  createdDateTime: string;
  id?: string;
  '@removed'?: {
    reason: 'changed' | 'deleted';
  };
}

export interface DeltaResponse<T> {
  '@odata.nextLink'?: string;
  '@odata.deltaLink'?: string;
  value: T[];
}

export class DeltaSyncError extends Error {
  constructor(
    message: string,
    public readonly code: string,
    public readonly statusCode: number,
  ) {
    super(message);
    this.name = 'DeltaSyncError';
  }
}

@Injectable()
export class DeltaSyncService {
  private readonly logger = new Logger(DeltaSyncService.name);
  private readonly MAX_RETRIES = 3;
  private readonly RETRY_DELAY_MS = 1000; // 1 second

  constructor(
    private readonly deltaLinkRepository: OutlookDeltaLinkRepository,
  ) {}

  private async delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  private async retryWithBackoff<T>(
    operation: () => Promise<T>,
    retryCount = 0,
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
    resourceType: ResourceType,
  ): void {
    // Handle sync reset (410 Gone)
    if (response['@odata.deltaLink']?.includes('$deltatoken=')) {
      this.logger.log(`Sync reset detected for user ${userId}, resource ${resourceType}`);
      // Clear the delta link to force a full sync
      this.deltaLinkRepository.saveDeltaLink(userId, resourceType, '');
    }

    // Handle token expiration
    if (response['@odata.deltaLink']) {
      const tokenExpiry = this.calculateTokenExpiry(resourceType);
      this.logger.log(`Delta token will expire at ${tokenExpiry.toISOString()}`);
    }
  }

  private calculateTokenExpiry(resourceType: ResourceType): Date {
    const now = new Date();
    // Directory objects and education objects have 7-day expiry
    if (resourceType === ResourceType.CALENDAR || resourceType === ResourceType.EMAIL) {
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
        if (item['@removed']) {
          uniqueItems.set(item.id, item);
        } 
        // If item exists and is not removed, update it
        else if (!uniqueItems.has(item.id) || !uniqueItems.get(item.id)?.['@removed']) {
          uniqueItems.set(item.id, item);
        }
      }
    }

    return Array.from(uniqueItems.values());
  }

  async fetchAndSortChanges<T extends DeltaItem>(
    userId: number,
    resourceType: ResourceType,
    accessToken: string,
    initialEndpoint: string,
  ): Promise<T[]> {
    try {
      const client = Client.init({
        authProvider: (done) => {
          done(null, accessToken);
        },
      });

      // Get the stored delta link for this user
      const deltaLink = await this.deltaLinkRepository.getDeltaLink(userId, resourceType);

      let requestUrl = initialEndpoint;

      // If we have a delta link, use that
      if (deltaLink) {
        requestUrl = deltaLink;
      }

      const allItems: T[] = [];

      // Fetch all pages of changes with retry logic
      let response: DeltaResponse<T> = { '@odata.nextLink': requestUrl, value: [] };

      while (response['@odata.nextLink']) {
        const nextLink = response['@odata.nextLink'];
        
        // Use retry logic for API calls
        response = await this.retryWithBackoff(async () => {
          try {
            return await client.api(nextLink).get();
          } catch (error: any) {
            // Handle token expiration
            if (error.statusCode === 401 || error.statusCode === 403) {
              throw new DeltaSyncError(
                'Token expired or invalid',
                'TokenExpired',
                error.statusCode
              );
            }
            // Handle sync reset
            if (error.statusCode === 410) {
              throw new DeltaSyncError(
                'Sync reset required',
                'SyncReset',
                error.statusCode
              );
            }
            throw error;
          }
        });

        if (response.value && Array.isArray(response.value)) {
          allItems.push(...response.value);
        }

        // Handle delta response (sync reset, token expiry)
        this.handleDeltaResponse(response, userId, resourceType);

        // Save the delta link if present
        if (response['@odata.deltaLink']) {
          await this.deltaLinkRepository.saveDeltaLink(
            userId,
            resourceType,
            response['@odata.deltaLink']
          );
        }
      }

      // Handle replays by deduplicating items
      const uniqueItems = this.handleReplays(allItems);

      // Sort the items by lastModifiedDateTime (or createdDateTime as fallback)
      const sortedItems = uniqueItems.sort((a, b) => {
        const aTime = a.lastModifiedDateTime || a.createdDateTime || '';
        const bTime = b.lastModifiedDateTime || b.createdDateTime || '';

        return new Date(bTime).getTime() - new Date(aTime).getTime();
      });

      this.logger.log(`Fetched and sorted ${sortedItems.length} ${resourceType} changes for user ${userId}`);

      return sortedItems;
    } catch (error: unknown) {
      if (error instanceof DeltaSyncError) {
        this.logger.error(`Delta sync error: ${error.message}`, error);
        throw error;
      }
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to fetch ${resourceType} changes: ${errorMessage}`);
      throw new Error(`Failed to fetch ${resourceType} changes: ${errorMessage}`);
    }
  }
} 