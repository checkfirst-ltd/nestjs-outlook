/**
 * Delay execution for a specified number of milliseconds
 * @param ms - Milliseconds to delay
 * @returns Promise that resolves after the delay
 */
export async function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * Retry an operation with exponential backoff
 * @param operation - The async operation to retry
 * @param options - Retry configuration options
 * @param options.maxRetries - Maximum number of retries (default: 3)
 * @param options.retryDelayMs - Base delay in milliseconds (default: 1000)
 * @param options.retryCount - Current retry count (used internally for recursion)
 * @returns The result of the operation
 * @throws The last error if all retries are exhausted
 */
export async function retryWithBackoff<T>(
  operation: () => Promise<T>,
  options?: {
    maxRetries?: number;
    retryDelayMs?: number;
    retryCount?: number;
  }
): Promise<T> {
  const maxRetries = options?.maxRetries ?? 3;
  const retryDelayMs = options?.retryDelayMs ?? 1000;
  const retryCount = options?.retryCount ?? 0;

  try {
    return await operation();
  } catch (error) {
    if (retryCount >= maxRetries) {
      throw error;
    }

    // Calculate exponential backoff delay
    const delayMs = retryDelayMs * Math.pow(2, retryCount);
    await delay(delayMs);

    return retryWithBackoff(operation, {
      maxRetries,
      retryDelayMs,
      retryCount: retryCount + 1,
    });
  }
}
