import {
  is404Error,
  is429Error,
  is503Error,
  isMailboxConcurrencyError,
  extractRetryAfterSeconds,
  retryWithBackoff,
} from './retry.util';
import { GraphRateLimiterService } from '../services/shared/graph-rate-limiter.service';

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
  /** Rate limiter service for per-user request throttling */
  rateLimiter?: GraphRateLimiterService;
  /** User ID for rate limiting (required if rateLimiter is provided) */
  userId?: string;
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
    maxRetries = 10,
    retryDelayMs = 1000,
    logger = {
      warn: (message: string) => { console.warn(message); },
      error: (message: string) => { console.error(message); },
    },
    resourceName = 'resource',
    return404AsNull = false,
    rateLimiter,
    userId,
  } = options;

  try {
    // Use retryWithBackoff utility with custom 429 handling
    return await retryWithBackoff(
      async () => {
        // Acquire rate limit permit before making request
        if (rateLimiter && userId) {
          await rateLimiter.acquirePermit(userId);
        }

        // Then claim a per-mailbox CONCURRENCY slot so no more than N requests are in flight
        // against this mailbox at once — the axis Outlook's MailboxConcurrency limit polices,
        // which the per-second rate window alone does not bound. Released in `finally` below.
        // Guarded by `typeof` so older/stubbed rate limiters without the method degrade to
        // rate-only behaviour instead of throwing.
        let mailboxSlot: string | null = null;
        if (rateLimiter && userId && typeof rateLimiter.acquireMailboxSlot === 'function') {
          mailboxSlot = await rateLimiter.acquireMailboxSlot(userId);
        }

        try {
          const result = await operation();

          // Notify circuit breaker of successful request
          if (rateLimiter) {
            await rateLimiter.recordSuccess();
          }

          return result;
        } catch (error) {
          // Special handling for 429 rate limit errors with Retry-After header
          // This takes precedence over standard exponential backoff
          if (is429Error(error)) {
            const retryAfterSeconds = extractRetryAfterSeconds(error);

            // An app-wide MailboxConcurrency throttle feeds the circuit breaker so the fleet sheds
            // load — the per-user cooldown below only slows one mailbox. Guarded for stubs.
            if (
              rateLimiter &&
              isMailboxConcurrencyError(error) &&
              typeof rateLimiter.recordThrottleFailure === 'function'
            ) {
              await rateLimiter.recordThrottleFailure();
            }

            if (retryAfterSeconds !== null) {
              // Notify rate limiter about 429 response (per-user cooldown). The actual wait is
              // performed once by retryWithBackoff, which also honours this Retry-After — do NOT
              // sleep here as well, or the backoff would be applied twice.
              if (rateLimiter && userId) {
                await rateLimiter.handleRateLimitResponse(
                  userId,
                  retryAfterSeconds,
                );
              }

              logger.warn(
                `Rate limited on ${resourceName}, backing off ${retryAfterSeconds}s as per Retry-After header`
              );
            }
          }

          // Special handling for 503 Service Unavailable
          if (is503Error(error)) {
            // Notify circuit breaker about 503 failure. The Retry-After wait itself is handled by
            // retryWithBackoff (single source of truth) — no local sleep here.
            if (rateLimiter) {
              await rateLimiter.record503Failure();
            }

            const retryAfterSeconds = extractRetryAfterSeconds(error);
            if (retryAfterSeconds !== null) {
              logger.warn(
                `Service unavailable (503) on ${resourceName}, backing off ${retryAfterSeconds}s as per Retry-After header`
              );
            }
          }

          // Re-throw to let retryWithBackoff handle retry logic
          throw error;
        } finally {
          // Always release the mailbox concurrency slot — including before a retry backoff, so a
          // waiting request isn't blocked by one that's sleeping. The next attempt re-acquires.
          if (rateLimiter && userId && mailboxSlot && typeof rateLimiter.releaseMailboxSlot === 'function') {
            await rateLimiter.releaseMailboxSlot(userId, mailboxSlot);
          }
        }
      },
      {
        maxRetries,
        retryDelayMs,
        logger: {
          warn: (message: string, context?: Record<string, unknown>) => {
            logger.warn(message, context);
          },
        },
        operationName: resourceName,
      }
    );
  } catch (error) {
    // Handle 404 - resource not found (likely deleted between webhook and this call)
    if (is404Error(error)) {
      if (return404AsNull) {
        logger.warn(`Resource not found (likely deleted): ${resourceName}`);
        return null;
      }
      logger.error(`Resource not found: ${resourceName}`, error);
    }

    throw error;
  }
}
