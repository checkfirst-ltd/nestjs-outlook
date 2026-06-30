import { RedisLike } from "../state/redis-like.interface";

/**
 * Optional shared-state backend configuration.
 *
 * When `redis.client` is provided, lock store and rate-limit store use Redis
 * so multiple containers coordinate. Otherwise both fall back to in-process
 * in-memory implementations (suitable only for single-container deployments).
 */
export interface MicrosoftOutlookStateConfig {
  redis?: {
    /** Host-provided ioredis-compatible client. nestjs-outlook never imports ioredis. */
    client: RedisLike;
    /** Key prefix applied to all keys. Defaults to "outlook:". */
    keyPrefix?: string;
    /**
     * If true, module init throws when the PING probe fails. Use in production
     * so a Redis outage triggers ECS restart instead of silently falling back
     * to in-memory (which re-introduces the cross-container concurrency bug).
     * Defaults to false.
     */
    required?: boolean;
  };
}

/**
 * Configuration interface for Microsoft Outlook OAuth settings
 */
export interface MicrosoftOutlookConfig {
  /**
   * The client id for the Microsoft Outlook OAuth settings
   */
  clientId: string;
  /**
   * The client secret for the Microsoft Outlook OAuth settings
   */
  clientSecret: string;
  /**
   * The path of the redirect uri. e.g. auth/microsoft/callback
   */
  redirectPath: string;
  /**
   * The base url of the backend. e.g. https://dev.dashboard.checkfirstapp.com
   */
  backendBaseUrl: string;
  /**
   * The base path of the backend. e.g. api/v1
   */
  basePath?: string;
  /**
   * The path for the calendar webhook endpoint. e.g. /calendar/webhook
   * Defaults to /calendar/webhook
   */
  calendarWebhookPath?: string;
  /**
   * Optional shared-state backend. Use Redis to coordinate locks and
   * rate-limit budgets across multiple containers. Omit for in-memory.
   */
  state?: MicrosoftOutlookStateConfig;
  /**
   * TTL (ms) for the "revocation email already sent this cycle" dedupe flag,
   * which suppresses duplicate USER_REFRESH_TOKEN_INVALID emits when a burst of
   * webhooks arrives for a newly-revoked user.
   *
   * The flag is normally cleared when the user re-authenticates (account becomes
   * ACTIVE again). This TTL only bounds the case where the user never reconnects,
   * so the flag self-heals instead of living forever. Defaults to one week
   * (7 * 24 * 60 * 60 * 1000 ms). Lower it to re-notify sooner, or raise it to
   * suppress duplicate emails for longer.
   */
  revocationEmitFlagTtlMs?: number;
}
