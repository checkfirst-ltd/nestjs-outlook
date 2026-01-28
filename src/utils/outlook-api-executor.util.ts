import {
  delay,
  is404Error,
  is429Error,
  isNetworkError,
  isNonRetryableError,
  isServerError,
  extractRetryAfterSeconds,
} from './retry.util';

/**
 * Options for executing Microsoft Graph API calls
 */
export interface GraphApiExecutorOptions {
  /** Maximum number of retry attempts (default: 3) */
  maxRetries?: number;
  /** Base delay in milliseconds for exponential backoff (default: 1000) */
  retryDelayMs?: number;
  /** Optional logger for debug/warn messages */
  logger?: {
    warn: (message: string, ...args: unknown[]) => void;
    error: (message: string, ...args: unknown[]) => void;
    debug?: (message: string, ...args: unknown[]) => void;
  };
  /** Optional resource name for logging context (e.g., "me/events/abc123") */
  resourceName?: string;
  /** If true, returns null on 404 instead of throwing (default: false) */
  return404AsNull?: boolean;
}

/**
 * Execute a Microsoft Graph API call with automatic retry and error handling
 *
 * Implements Microsoft Graph best practices:
 * - Respects Retry-After header on 429 rate limit errors (fastest recovery)
 * - Uses exponential backoff for transient failures
 * - Does NOT retry permanent client errors (401, 403, 404)
 * - Automatically retries network errors and server errors (5xx)
 *
 * @param operation - The async operation to execute (should return a Promise)
 * @param options - Configuration options
 * @returns The result of the operation, or null if 404 and return404AsNull is true
 * @throws Error if non-retryable error or max retries exceeded
 *
 * @example
 * // Basic usage - fetch event details
 * const event = await executeGraphApiCall(
 *   () => client.api('/me/events/123').get(),
 *   { logger, resourceName: 'me/events/123' }
 * );
 *
 * @example
 * // Return null on 404 (common for deleted events)
 * const event = await executeGraphApiCall(
 *   () => client.api('/me/events/123').get(),
 *   { return404AsNull: true, logger }
 * );
 * // event will be null if the resource was deleted
 *
 * @example
 * // Custom retry configuration
 * const result = await executeGraphApiCall(
 *   () => someGraphApiCall(),
 *   { maxRetries: 5, retryDelayMs: 2000 }
 * );
 */
export async function executeGraphApiCall<T>(
  operation: () => Promise<T>,
  options: GraphApiExecutorOptions = {}
): Promise<T | null> {
  const {
    maxRetries = 3,
    retryDelayMs = 1000,
    logger = {
      warn: (message: string) => { console.warn(message); },
      error: (message: string) => { console.error(message); },
    },
    resourceName = 'resource',
    return404AsNull = false,
  } = options;

  for (let retryCount = 0; retryCount < maxRetries; retryCount++) {
    try {
      return await operation();
    } catch (error) {
      // Handle 404 - resource not found (likely deleted between webhook and this call)
      if (is404Error(error)) {
        if (return404AsNull) {
          logger.warn(`Resource not found (likely deleted): ${resourceName}`);
          return null;
        }
        // Otherwise throw immediately - 404 is non-retryable
        logger.error(`Resource not found: ${resourceName}`, error);
        throw error;
      }

      // Handle 401 and other non-retryable errors - throw immediately
      if (isNonRetryableError(error)) {
        logger.error(`Non-retryable error for ${resourceName}`, error);
        throw error;
      }

      // Check if we've exhausted retries
      if (retryCount >= maxRetries) {
        logger.error(`Max retries (${maxRetries}) exceeded for ${resourceName}`, error);
        throw error;
      }

      // Handle 429 - rate limit with Retry-After header (Microsoft's recommended fastest recovery)
      if (is429Error(error)) {
        const retryAfterSeconds = extractRetryAfterSeconds(error);
        const delayMs = retryAfterSeconds !== null
          ? retryAfterSeconds * 1000
          : retryDelayMs * Math.pow(2, retryCount); // Fallback to exponential backoff

        logger.warn(
          `Rate limited on ${resourceName}, retrying after ${delayMs / 1000}s (${maxRetries - retryCount} attempts remaining)`
        );

        await delay(delayMs);
        retryCount++;
        continue;
      }

      // Handle network/timeout errors with exponential backoff
      if (isNetworkError(error)) {
        const delayMs = retryDelayMs * Math.pow(2, retryCount);
        logger.warn(
          `Network timeout on ${resourceName}, retrying after ${delayMs}ms (${maxRetries - retryCount} attempts remaining)`
        );

        await delay(delayMs);
        retryCount++;
        continue;
      }

      // Handle server errors (5xx) with exponential backoff
      if (isServerError(error)) {
        const delayMs = retryDelayMs * Math.pow(2, retryCount);
        logger.warn(
          `Server error on ${resourceName}, retrying after ${delayMs}ms (${maxRetries - retryCount} attempts remaining)`
        );

        await delay(delayMs);
        retryCount++;
        continue;
      }

      // Unknown error - throw immediately
      logger.error(`Unknown error for ${resourceName}`, error);
      throw error;
    }
  }

  throw new Error(`Max retries (${maxRetries}) exceeded for ${resourceName}`);
}
