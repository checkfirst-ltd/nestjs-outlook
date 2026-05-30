import { Inject, Injectable, Logger, OnModuleInit } from '@nestjs/common';
import { Cron, CronExpression } from '@nestjs/schedule';
import { delay } from '../../utils/retry.util';
import {
  OUTLOOK_RATE_LIMIT_STORE,
} from '../../constants';
import {
  OutlookRateLimitStore,
} from './outlook-rate-limit.store';

/**
 * Global per-user rate limiter for Microsoft Graph API
 *
 * Implements Microsoft's rate limits:
 * - 4 requests per second per user (mailbox)
 * - 10,000 requests per 10 minutes per user
 *
 * State is delegated to an OutlookRateLimitStore — in-memory by default,
 * Redis when configured. The Redis backend coordinates budgets and the
 * circuit breaker across all containers in the fleet.
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

  // Cleanup configuration (in-memory backend only)
  private readonly INACTIVE_USER_THRESHOLD_MS = 30 * 60 * 1000; // 30 minutes

  // Stats for monitoring
  private totalPermitsAcquired = 0;
  private totalWaitTime = 0;
  private cooldownCount = 0;

  // Circuit breaker configuration (service-level, not per-user)
  private readonly CB_FAILURE_THRESHOLD = 5;       // consecutive 503s to trip
  private readonly CB_FAILURE_WINDOW_MS = 60_000;  // 503s must occur within 60s
  private readonly CB_COOLDOWN_MS = 60_000;        // how long to stay open
  private readonly CB_PROBE_TTL_MS = 5_000;        // half-open probe slot TTL

  // In-process failure timestamp tracking (mirrors store CB state for triggering)
  private cbFailureTimestamps: number[] = [];
  private cbTotalTrips = 0;

  constructor(
    @Inject(OUTLOOK_RATE_LIMIT_STORE)
    private readonly store: OutlookRateLimitStore,
  ) {}

  onModuleInit() {
    this.logger.log(
      `GraphRateLimiterService initialized (backend=${this.store.kind}) - Limits: ${this.MAX_REQUESTS_PER_SECOND} req/sec, ` +
      `${this.MAX_REQUESTS_PER_10_MINUTES} req/10min per user`
    );
  }

  /**
   * Acquire permission to make a Graph API request for a user.
   * Blocks until the request can be made within rate limits.
   *
   * @param userId - External user ID
   */
  async acquirePermit(userId: string): Promise<void> {
    await this.waitForCircuitBreaker();

    const startTime = Date.now();
    let waitedOnce = false;

    while (true) {
      const cooldownUntil = await this.store.getCooldown(userId);
      if (cooldownUntil && cooldownUntil > Date.now()) {
        const wait = cooldownUntil - Date.now() + 50;
        if (!waitedOnce) {
          waitedOnce = true;
          this.logger.debug(
            `[acquirePermit] User ${userId} in cooldown for ${wait}ms`,
          );
        }
        await delay(Math.max(wait, 50));
        continue;
      }

      const secCount = await this.store.getCount(
        userId,
        this.ONE_SECOND_MS,
        'sec',
      );
      const tenMinCount = await this.store.getCount(
        userId,
        this.TEN_MINUTES_MS,
        'min10',
      );

      if (secCount >= this.MAX_REQUESTS_PER_SECOND) {
        if (!waitedOnce) {
          waitedOnce = true;
          this.logger.debug(
            `[acquirePermit] User ${userId} at 1s limit, waiting`,
          );
        }
        await delay(100);
        continue;
      }

      if (tenMinCount >= this.MAX_REQUESTS_PER_10_MINUTES) {
        this.logger.warn(
          `[acquirePermit] User ${userId} hit 10-minute limit (${tenMinCount}/${this.MAX_REQUESTS_PER_10_MINUTES})`,
        );
        await delay(1000);
        continue;
      }

      break;
    }

    // Record the request atomically in both windows.
    await this.store.recordRequest(userId, this.ONE_SECOND_MS, 'sec');
    await this.store.recordRequest(userId, this.TEN_MINUTES_MS, 'min10');

    this.totalPermitsAcquired++;
    const totalWait = Date.now() - startTime;
    if (totalWait > 50) {
      this.totalWaitTime += totalWait;
      this.logger.debug(
        `[acquirePermit] User ${userId} waited ${totalWait}ms for permit`,
      );
    }
  }

  /**
   * Release a permit after request completes.
   * No-op since permits are tracked on the acquire side.
   */
  releasePermit(_userId: string): void {
    // No-op - we track on acquire side.
  }

  /**
   * Handle a 429 rate limit response from Microsoft Graph API.
   * Sets a cooldown period for the user based on Retry-After header.
   */
  async handleRateLimitResponse(
    userId: string,
    retryAfterSeconds: number,
  ): Promise<void> {
    const cooldownUntil = Date.now() + retryAfterSeconds * 1000;
    await this.store.setCooldown(userId, cooldownUntil);
    this.cooldownCount++;
    this.logger.warn(
      `[handleRateLimitResponse] User ${userId} hit 429 - cooldown for ${retryAfterSeconds}s ` +
      `until ${new Date(cooldownUntil).toISOString()}`,
    );
  }

  /**
   * Cleanup inactive users (in-memory backend only).
   * Redis backend uses key TTLs.
   */
  @Cron(CronExpression.EVERY_5_MINUTES)
  private async cleanupInactiveUsers(): Promise<void> {
    if (this.store.kind !== 'memory') return;

    const removed = await this.store.cleanupInactive(
      this.INACTIVE_USER_THRESHOLD_MS,
    );

    const now = Date.now();
    this.cbFailureTimestamps = this.cbFailureTimestamps.filter(
      (t) => now - t < this.CB_FAILURE_WINDOW_MS,
    );

    if (removed > 0) {
      this.logger.log(
        `[cleanupInactiveUsers] Cleaned up ${removed} inactive users`,
      );
    }
  }

  /**
   * Record a 503 Service Unavailable failure (service-level circuit breaker).
   * When consecutive failures exceed threshold within the window, circuit opens.
   */
  async record503Failure(): Promise<void> {
    const now = Date.now();
    this.cbFailureTimestamps.push(now);
    this.cbFailureTimestamps = this.cbFailureTimestamps.filter(
      (t) => now - t < this.CB_FAILURE_WINDOW_MS,
    );

    const cb = await this.store.getCbState();
    const state = cb?.state ?? 'closed';

    if (state === 'half-open') {
      await this.store.setCbState({ state: 'open', openedAt: now });
      this.logger.warn(
        `[CircuitBreaker] Half-open probe failed (503), re-opening circuit for ${this.CB_COOLDOWN_MS / 1000}s`,
      );
      return;
    }

    if (
      state === 'closed' &&
      this.cbFailureTimestamps.length >= this.CB_FAILURE_THRESHOLD
    ) {
      await this.store.setCbState({ state: 'open', openedAt: now });
      this.cbTotalTrips++;
      this.logger.warn(
        `[CircuitBreaker] OPEN — ${this.cbFailureTimestamps.length} 503 errors ` +
        `within ${this.CB_FAILURE_WINDOW_MS / 1000}s window. ` +
        `Blocking requests for ${this.CB_COOLDOWN_MS / 1000}s`,
      );
    }
  }

  /**
   * Record a successful Graph API response (resets circuit breaker if half-open).
   */
  async recordSuccess(): Promise<void> {
    const cb = await this.store.getCbState();
    if (cb?.state === 'half-open') {
      await this.store.setCbState({ state: 'closed', openedAt: null });
      this.cbFailureTimestamps = [];
      this.logger.log('[CircuitBreaker] CLOSED — half-open probe succeeded');
      return;
    }
    if (cb?.state === 'closed') {
      this.cbFailureTimestamps = [];
    }
  }

  /**
   * Check circuit breaker state and wait if open.
   * Transitions open -> half-open after cooldown expires.
   */
  private async waitForCircuitBreaker(): Promise<void> {
    const cb = await this.store.getCbState();
    if (!cb || cb.state === 'closed') return;

    const now = Date.now();

    if (cb.state === 'open') {
      const elapsed = now - (cb.openedAt ?? now);
      if (elapsed >= this.CB_COOLDOWN_MS) {
        const wonProbe = await this.store.tryClaimHalfOpenProbe(
          this.CB_PROBE_TTL_MS,
        );
        if (wonProbe) {
          await this.store.setCbState({ state: 'half-open', openedAt: null });
          this.logger.log(
            '[CircuitBreaker] HALF-OPEN — cooldown expired, this container is probing',
          );
          return;
        }
        // Another container is probing; wait one probe TTL and retry.
        await delay(this.CB_PROBE_TTL_MS);
        return this.waitForCircuitBreaker();
      }

      const remaining = this.CB_COOLDOWN_MS - elapsed;
      this.logger.warn(
        `[CircuitBreaker] Circuit OPEN, waiting ${Math.round(remaining / 1000)}s before retry`,
      );
      await delay(remaining);
      return this.waitForCircuitBreaker();
    }

    if (cb.state === 'half-open') {
      // Only the container that won the probe slot proceeds; others wait.
      const wonProbe = await this.store.tryClaimHalfOpenProbe(
        this.CB_PROBE_TTL_MS,
      );
      if (!wonProbe) {
        await delay(this.CB_PROBE_TTL_MS);
        return this.waitForCircuitBreaker();
      }
    }
  }

  /**
   * Get rate limiter statistics for monitoring.
   */
  async getStats() {
    const cb = await this.store.getCbState();
    return {
      backend: this.store.kind,
      activeUsers: await this.store.getActiveUserCount(),
      totalPermitsAcquired: this.totalPermitsAcquired,
      totalWaitTimeMs: this.totalWaitTime,
      averageWaitTimeMs:
        this.totalPermitsAcquired > 0
          ? Math.round(this.totalWaitTime / this.totalPermitsAcquired)
          : 0,
      cooldownCount: this.cooldownCount,
      circuitBreaker: {
        state: cb?.state ?? 'closed',
        recentFailures: this.cbFailureTimestamps.length,
        totalTrips: this.cbTotalTrips,
        openedAt: cb?.openedAt ? new Date(cb.openedAt).toISOString() : null,
      },
    };
  }

  /**
   * Get per-user statistics (for debugging).
   */
  async getUserStats(userId: string) {
    const [secCount, tenMinCount, cooldownUntil] = await Promise.all([
      this.store.getCount(userId, this.ONE_SECOND_MS, 'sec'),
      this.store.getCount(userId, this.TEN_MINUTES_MS, 'min10'),
      this.store.getCooldown(userId),
    ]);
    return {
      userId,
      backend: this.store.kind,
      recentRequestCount: secCount,
      tenMinuteRequestCount: tenMinCount,
      cooldownUntil,
    };
  }
}
