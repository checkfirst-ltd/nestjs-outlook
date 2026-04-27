import { Injectable, Logger, OnModuleInit } from '@nestjs/common';
import { Cron, CronExpression } from '@nestjs/schedule';
import { delay } from '../../utils/retry.util';

/**
 * Per-user rate limiter state
 */
interface UserRateLimiter {
  userId: string;
  // Sliding window for 1-second limit (4 requests/sec)
  recentRequests: number[]; // timestamps in ms
  // Sliding window for 10-minute limit (10,000 requests)
  tenMinuteWindow: number[]; // timestamps in ms
  // Cooldown until timestamp (from Retry-After header)
  cooldownUntil: number | null;
  // Last activity timestamp for cleanup
  lastActivity: number;
}

/**
 * Global per-user rate limiter for Microsoft Graph API
 *
 * Implements Microsoft's rate limits:
 * - 4 requests per second per user (mailbox)
 * - 10,000 requests per 10 minutes per user
 *
 * Features:
 * - Sliding window algorithm for accurate quota tracking
 * - Per-user request queueing
 * - Retry-After header support
 * - Automatic cooldown management
 * - Inactive user cleanup
 *
 * @see https://learn.microsoft.com/en-us/graph/throttling-limits
 */
@Injectable()
export class GraphRateLimiterService implements OnModuleInit {
  private readonly logger = new Logger(GraphRateLimiterService.name);

  // Rate limit thresholds (per user)
  private readonly MAX_REQUESTS_PER_SECOND = 4;
  private readonly MAX_REQUESTS_PER_10_MINUTES = 10000;
  private readonly ONE_SECOND_MS = 1000;
  private readonly TEN_MINUTES_MS = 10 * 60 * 1000;

  // Cleanup configuration
  private readonly INACTIVE_USER_THRESHOLD_MS = 30 * 60 * 1000; // 30 minutes

  // Per-user rate limiters
  private readonly userLimiters = new Map<string, UserRateLimiter>();

  // Stats for monitoring
  private totalPermitsAcquired = 0;
  private totalWaitTime = 0;
  private cooldownCount = 0;

  // Circuit breaker configuration (service-level, not per-user)
  private readonly CB_FAILURE_THRESHOLD = 5;       // consecutive 503s to trip
  private readonly CB_FAILURE_WINDOW_MS = 60_000;  // 503s must occur within 60s
  private readonly CB_COOLDOWN_MS = 60_000;        // how long to stay open
  private readonly CB_HALF_OPEN_MAX = 1;           // requests allowed in half-open

  // Circuit breaker state
  private cbState: 'closed' | 'open' | 'half-open' = 'closed';
  private cbFailureTimestamps: number[] = [];
  private cbOpenedAt: number | null = null;
  private cbHalfOpenInFlight = 0;
  private cbTotalTrips = 0;

  onModuleInit() {
    this.logger.log(
      `GraphRateLimiterService initialized - Limits: ${this.MAX_REQUESTS_PER_SECOND} req/sec, ` +
      `${this.MAX_REQUESTS_PER_10_MINUTES} req/10min per user`
    );
  }

  /**
   * Acquire permission to make a Graph API request for a user
   * Blocks until the request can be made within rate limits
   *
   * @param userId - External user ID
   */
  async acquirePermit(userId: string): Promise<void> {
    // Service-level circuit breaker check (blocks if Microsoft Graph is down)
    await this.waitForCircuitBreaker();

    // Initialize limiter if needed
    if (!this.userLimiters.has(userId)) {
      this.initializeLimiter(userId);
    }

    const limiter = this.userLimiters.get(userId);
    if (!limiter) {
      this.logger.warn(
        `[acquirePermit] No limiter found for user ${userId}, skipping acquire`
      );
      return;
    }

    const startTime = Date.now();
    let waitedOnce = false;

    // Wait until we can make request
    while (!this.canMakeRequest(userId)) {
      if (!waitedOnce) {
        waitedOnce = true;
        this.logger.debug(
          `[acquirePermit] User ${userId} rate limited, entering wait loop`
        );
      }

      // Calculate wait time based on oldest request in 1-sec window
      const oldestRequest = limiter.recentRequests[0];
      const waitTime = oldestRequest
        ? this.ONE_SECOND_MS - (Date.now() - oldestRequest) + 50
        : 100;

      await delay(Math.max(waitTime, 50)); // Min 50ms wait
    }

    // Record this request
    const now = Date.now();
    limiter.recentRequests.push(now);
    limiter.tenMinuteWindow.push(now);
    limiter.lastActivity = now;

    // Track stats
    this.totalPermitsAcquired++;
    const totalWait = now - startTime;
    if (totalWait > 50) {
      this.totalWaitTime += totalWait;
      this.logger.debug(
        `[acquirePermit] User ${userId} waited ${totalWait}ms for permit ` +
        `(recent: ${limiter.recentRequests.length}/${this.MAX_REQUESTS_PER_SECOND}, ` +
        `10min: ${limiter.tenMinuteWindow.length}/${this.MAX_REQUESTS_PER_10_MINUTES})`
      );
    }
  }

  /**
   * Release a permit after request completes
   * Currently a no-op since we track permits on acquire, but provided for symmetry
   *
   * @param userId - External user ID
   */
  releasePermit(userId: string): void {
    if (!this.userLimiters.has(userId)) {
      this.logger.warn(
        `[releasePermit] No limiter found for user ${userId}, skipping release`
      );
      return;
    }

    // No-op - we track on acquire side
    // Kept for API symmetry and potential future use
  }

  /**
   * Handle a 429 rate limit response from Microsoft Graph API
   * Sets a cooldown period for the user based on Retry-After header
   *
   * @param userId - External user ID
   * @param retryAfterSeconds - Seconds to wait from Retry-After header
   */
  handleRateLimitResponse(userId: string, retryAfterSeconds: number): void {
    if (!this.userLimiters.has(userId)) {
      this.initializeLimiter(userId);
    }

    const limiter = this.userLimiters.get(userId);
    if (!limiter) {
      this.logger.warn(
        `[handleRateLimitResponse] No limiter found for user ${userId}, skipping cooldown`
      );
      return;
    }

    const cooldownUntil = Date.now() + (retryAfterSeconds * 1000);
    limiter.cooldownUntil = cooldownUntil;

    this.cooldownCount++;

    this.logger.warn(
      `[handleRateLimitResponse] User ${userId} hit 429 - cooldown for ${retryAfterSeconds}s ` +
      `until ${new Date(cooldownUntil).toISOString()}`
    );
  }

  /**
   * Check if a user can make a request within rate limits
   *
   * @param userId - External user ID
   * @returns True if request can be made, false if rate limited
   */
  private canMakeRequest(userId: string): boolean {
    const limiter = this.userLimiters.get(userId);
    if (!limiter) return true;

    const now = Date.now();

    // Clean up old requests from 1-second window
    limiter.recentRequests = limiter.recentRequests.filter(
      timestamp => now - timestamp < this.ONE_SECOND_MS
    );

    // Clean up old requests from 10-minute window
    limiter.tenMinuteWindow = limiter.tenMinuteWindow.filter(
      timestamp => now - timestamp < this.TEN_MINUTES_MS
    );

    // Check if under cooldown from 429 response
    if (limiter.cooldownUntil && now < limiter.cooldownUntil) {
      return false;
    }

    // Check 1-second rate limit
    if (limiter.recentRequests.length >= this.MAX_REQUESTS_PER_SECOND) {
      return false;
    }

    // Check 10-minute rate limit
    if (limiter.tenMinuteWindow.length >= this.MAX_REQUESTS_PER_10_MINUTES) {
      this.logger.warn(
        `[canMakeRequest] User ${userId} hit 10-minute limit ` +
        `(${limiter.tenMinuteWindow.length}/${this.MAX_REQUESTS_PER_10_MINUTES})`
      );
      return false;
    }

    return true;
  }

  /**
   * Initialize rate limiter for a new user
   *
   * @param userId - External user ID
   */
  private initializeLimiter(userId: string): void {
    this.userLimiters.set(userId, {
      userId,
      recentRequests: [],
      tenMinuteWindow: [],
      cooldownUntil: null,
      lastActivity: Date.now(),
    });

    this.logger.debug(`[initializeLimiter] Created limiter for user ${userId}`);
  }

  /**
   * Cleanup inactive users (runs every 5 minutes via cron)
   * Removes rate limiter state for users inactive for >30 minutes
   */
  @Cron(CronExpression.EVERY_5_MINUTES)
  private cleanupInactiveUsers(): void {
    const now = Date.now();
    let cleaned = 0;

    for (const [userId, limiter] of this.userLimiters.entries()) {
      const inactiveDuration = now - limiter.lastActivity;

      if (inactiveDuration > this.INACTIVE_USER_THRESHOLD_MS) {
        this.userLimiters.delete(userId);
        cleaned++;
      }
    }

    // Prune stale circuit breaker failure timestamps
    this.cbFailureTimestamps = this.cbFailureTimestamps.filter(
      t => now - t < this.CB_FAILURE_WINDOW_MS
    );

    if (cleaned > 0) {
      this.logger.log(
        `[cleanupInactiveUsers] Cleaned up ${cleaned} inactive users ` +
        `(${this.userLimiters.size} active limiters remaining)`
      );
    }
  }

  /**
   * Record a 503 Service Unavailable failure (service-level circuit breaker).
   * When consecutive failures exceed threshold within the window, circuit opens.
   */
  record503Failure(): void {
    const now = Date.now();

    this.cbFailureTimestamps.push(now);
    this.cbFailureTimestamps = this.cbFailureTimestamps.filter(
      t => now - t < this.CB_FAILURE_WINDOW_MS
    );

    if (this.cbState === 'half-open') {
      // Probe failed — re-open circuit
      this.cbState = 'open';
      this.cbOpenedAt = now;
      this.cbHalfOpenInFlight = 0;
      this.logger.warn(
        `[CircuitBreaker] Half-open probe failed (503), re-opening circuit for ${this.CB_COOLDOWN_MS / 1000}s`
      );
      return;
    }

    if (
      this.cbState === 'closed' &&
      this.cbFailureTimestamps.length >= this.CB_FAILURE_THRESHOLD
    ) {
      this.cbState = 'open';
      this.cbOpenedAt = now;
      this.cbTotalTrips++;
      this.logger.warn(
        `[CircuitBreaker] OPEN — ${this.cbFailureTimestamps.length} 503 errors ` +
        `within ${this.CB_FAILURE_WINDOW_MS / 1000}s window. ` +
        `Blocking requests for ${this.CB_COOLDOWN_MS / 1000}s`
      );
    }
  }

  /**
   * Record a successful Graph API response (resets circuit breaker).
   */
  recordSuccess(): void {
    if (this.cbState === 'half-open') {
      this.logger.log('[CircuitBreaker] CLOSED — half-open probe succeeded');
      this.cbState = 'closed';
      this.cbFailureTimestamps = [];
      this.cbHalfOpenInFlight = 0;
      this.cbOpenedAt = null;
      return;
    }

    if (this.cbState === 'closed') {
      // Reset consecutive failure tracking on success
      this.cbFailureTimestamps = [];
    }
  }

  /**
   * Check circuit breaker state and wait if open.
   * Transitions open -> half-open after cooldown expires.
   */
  private async waitForCircuitBreaker(): Promise<void> {
    if (this.cbState === 'closed') return;

    const now = Date.now();

    if (this.cbState === 'open') {
      const elapsed = now - (this.cbOpenedAt ?? now);
      if (elapsed >= this.CB_COOLDOWN_MS) {
        this.cbState = 'half-open';
        this.cbHalfOpenInFlight = 0;
        this.logger.log('[CircuitBreaker] HALF-OPEN — cooldown expired, allowing probe request');
      } else {
        const remaining = this.CB_COOLDOWN_MS - elapsed;
        this.logger.warn(
          `[CircuitBreaker] Circuit OPEN, waiting ${Math.round(remaining / 1000)}s before retry`
        );
        await delay(remaining);
        this.cbState = 'half-open';
        this.cbHalfOpenInFlight = 0;
        this.logger.log('[CircuitBreaker] HALF-OPEN — cooldown expired after wait');
      }
    }

    // Re-read state (may have changed asynchronously via recordSuccess/record503Failure)
    const currentState = this.cbState as 'closed' | 'open' | 'half-open';
    if (currentState === 'half-open') {
      // Only allow CB_HALF_OPEN_MAX concurrent probe requests
      while (this.cbHalfOpenInFlight >= this.CB_HALF_OPEN_MAX) {
        await delay(500);
        // State may change asynchronously via recordSuccess/record503Failure
        const stateAfterWait: string = this.cbState;
        if (stateAfterWait === 'closed') return;
        if (stateAfterWait === 'open') {
          return this.waitForCircuitBreaker();
        }
      }
      this.cbHalfOpenInFlight++;
    }
  }

  /**
   * Get rate limiter statistics for monitoring
   */
  getStats() {
    return {
      activeUsers: this.userLimiters.size,
      totalPermitsAcquired: this.totalPermitsAcquired,
      totalWaitTimeMs: this.totalWaitTime,
      averageWaitTimeMs: this.totalPermitsAcquired > 0
        ? Math.round(this.totalWaitTime / this.totalPermitsAcquired)
        : 0,
      cooldownCount: this.cooldownCount,
      circuitBreaker: {
        state: this.cbState,
        recentFailures: this.cbFailureTimestamps.length,
        totalTrips: this.cbTotalTrips,
        openedAt: this.cbOpenedAt ? new Date(this.cbOpenedAt).toISOString() : null,
      },
    };
  }

  /**
   * Get per-user statistics (for debugging)
   */
  getUserStats(userId: string) {
    const limiter = this.userLimiters.get(userId);
    if (!limiter) {
      return null;
    }

    const now = Date.now();
    return {
      userId,
      recentRequestCount: limiter.recentRequests.filter(
        t => now - t < this.ONE_SECOND_MS
      ).length,
      tenMinuteRequestCount: limiter.tenMinuteWindow.filter(
        t => now - t < this.TEN_MINUTES_MS
      ).length,
      cooldownUntil: limiter.cooldownUntil,
      lastActivity: new Date(limiter.lastActivity).toISOString(),
    };
  }
}
