/**
 * Delay execution for a specified number of milliseconds
 * @param ms - Milliseconds to delay
 * @returns Promise that resolves after the delay
 */
export async function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * Check if an error is a Microsoft Graph API error with a specific status code
 * @param error - The error to check
 * @param statusCode - The HTTP status code to match
 * @returns True if the error matches the status code
 */
function isGraphErrorWithStatus(error: unknown, statusCode: number): boolean {
  if (!error || typeof error !== 'object') {
    return false;
  }

  // Check for Microsoft Graph SDK error format
  if ('statusCode' in error && typeof error.statusCode === 'number') {
    return error.statusCode === statusCode;
  }

  // Check for nested error in stack array (Microsoft Graph SDK format)
  if ('stack' in error && Array.isArray(error.stack) && error.stack.length > 0) {
    const firstError = error.stack[0];
    if (firstError && typeof firstError === 'object' && 'statusCode' in firstError) {
      return firstError.statusCode === statusCode;
    }
  }

  return false;
}

/**
 * Check if an error is non-retryable (permanent failure)
 * Non-retryable errors include:
 * - 401 Unauthorized (invalid/expired access token)
 * - 403 Forbidden (insufficient permissions)
 * - 404 Not Found (resource doesn't exist)
 * - 410 Gone (sync state/delta token expired)
 * @param error - The error to check
 * @returns True if the error should not be retried
 */
export function isNonRetryableError(error: unknown): boolean {
  return (
    isGraphErrorWithStatus(error, 401) ||
    isGraphErrorWithStatus(error, 403) ||
    isGraphErrorWithStatus(error, 404) ||
    isGraphErrorWithStatus(error, 410)
  );
}

/**
 * Check if an error is a 410 Gone error (sync state/delta token expired)
 * @param error - The error to check
 * @returns True if the error is a 410 Gone error
 */
export function is410Error(error: unknown): boolean {
  return isGraphErrorWithStatus(error, 410);
}

/**
 * Retry an operation with exponential backoff
 * @param operation - The async operation to retry
 * @param options - Retry configuration options
 * @param options.maxRetries - Maximum number of retries (default: 3)
 * @param options.retryDelayMs - Base delay in milliseconds (default: 1000)
 * @param options.retryCount - Current retry count (used internally for recursion)
 * @param options.logger - Optional logger for retry attempts
 * @param options.operationName - Optional name for logging context
 * @returns The result of the operation
 * @throws The last error if all retries are exhausted or if error is non-retryable
 */
export async function retryWithBackoff<T>(
  operation: () => Promise<T>,
  options?: {
    maxRetries?: number;
    retryDelayMs?: number;
    retryCount?: number;
    logger?: { warn: (message: string, context?: any) => void };
    operationName?: string;
  }
): Promise<T> {
  const maxRetries = options?.maxRetries ?? 3;
  const retryDelayMs = options?.retryDelayMs ?? 1000;
  const retryCount = options?.retryCount ?? 0;
  const logger = options?.logger;
  const operationName = options?.operationName ?? 'operation';

  try {
    return await operation();
  } catch (error) {
    // Extract error details for logging
    const errorDetails = extractErrorInfo(error);

    // Don't retry non-retryable errors (401, 403, 404, 410)
    if (isNonRetryableError(error)) {
      if (logger) {
        logger.warn(`[retryWithBackoff] Non-retryable error for ${operationName}, not retrying`, {
          statusCode: errorDetails.statusCode,
          errorCode: errorDetails.code,
        });
      }
      throw error;
    }

    if (retryCount >= maxRetries) {
      if (logger) {
        logger.warn(`[retryWithBackoff] Max retries (${maxRetries}) exceeded for ${operationName}`, {
          statusCode: errorDetails.statusCode,
          errorCode: errorDetails.code,
          errorMessage: errorDetails.message,
        });
      }
      throw error;
    }

    // Calculate exponential backoff delay
    const delayMs = retryDelayMs * Math.pow(2, retryCount);

    if (logger) {
      logger.warn(`[retryWithBackoff] Retry ${retryCount + 1}/${maxRetries} for ${operationName} after ${delayMs}ms`, {
        statusCode: errorDetails.statusCode,
        errorCode: errorDetails.code,
        errorType: errorDetails.type,
        delayMs,
      });
    }

    await delay(delayMs);

    return retryWithBackoff(operation, {
      maxRetries,
      retryDelayMs,
      retryCount: retryCount + 1,
      logger,
      operationName,
    });
  }
}

/**
 * Extract error information for logging
 * @param error - The error object
 * @returns Error details
 */
function extractErrorInfo(error: unknown): {
  statusCode: number | string;
  code: string;
  message: string;
  type: string;
} {
  if (!error || typeof error !== 'object') {
    return {
      statusCode: 'N/A',
      code: 'N/A',
      message: String(error),
      type: 'unknown',
    };
  }

  // Check for Microsoft Graph SDK error format
  if ('statusCode' in error) {
    const statusCode = error.statusCode as number;
    const code = 'code' in error ? String(error.code) : 'N/A';
    const body = 'body' in error ? String(error.body) : 'N/A';

    return {
      statusCode,
      code,
      message: body,
      type: statusCode === -1 ? 'network_error' : 'graph_api_error',
    };
  }

  // Check for nested error in stack array
  if ('stack' in error && Array.isArray(error.stack) && error.stack.length > 0) {
    const firstError = error.stack[0];
    if (firstError && typeof firstError === 'object') {
      const statusCode = 'statusCode' in firstError ? (firstError.statusCode as number) : 'N/A';
      const code = 'code' in firstError ? String(firstError.code) : 'N/A';
      const body = 'body' in firstError ? String(firstError.body) : 'N/A';

      return {
        statusCode,
        code,
        message: body,
        type: statusCode === -1 ? 'network_error' : 'graph_api_error',
      };
    }
  }

  return {
    statusCode: 'N/A',
    code: 'N/A',
    message: error instanceof Error ? error.message : String(error),
    type: 'generic_error',
  };
}
