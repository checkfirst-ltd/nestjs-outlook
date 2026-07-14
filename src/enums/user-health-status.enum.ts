/**
 * Connection-health verdict for a single user, combining the `microsoft_users` row, its active
 * Outlook webhook subscription, and (optionally) Microsoft Graph. Only `HEALTHY` means "connected".
 */
export enum UserHealthStatus {
  /** Active user with a live calendar subscription. Connected. */
  HEALTHY = 'HEALTHY',

  // ── recoverable (auto-fixable by recreating the subscription) ──
  /** Mapped/active user but no active calendar subscription. */
  NO_SUBSCRIPTION = 'NO_SUBSCRIPTION',
  /** Subscription exists but has passed its expiration. */
  SUBSCRIPTION_EXPIRED = 'SUBSCRIPTION_EXPIRED',
  /** Subscription active locally but no notification received within the stale window. */
  SUBSCRIPTION_STALE = 'SUBSCRIPTION_STALE',
  /** Subscription active locally but returns 404 at Microsoft (only when verifyAtGraph is set). */
  MISSING_AT_GRAPH = 'MISSING_AT_GRAPH',

  // ── not auto-recoverable (reported, needs a human) ──
  /** Delegated user whose token is dead (status CORRUPTED) — needs re-authentication. */
  NEEDS_REAUTH = 'NEEDS_REAUTH',
  /** App-only tenant is revoked / cert-expired / disabled — needs an administrator. */
  NEEDS_ADMIN = 'NEEDS_ADMIN',
  /** App-only user with no tenant mapping (cannot recover without an email/UPN here). */
  NOT_MAPPED = 'NOT_MAPPED',
  /** Row is soft-deleted (`isActive = false`). */
  INACTIVE = 'INACTIVE',
  /** No `microsoft_users` row exists for this external id. */
  UNKNOWN = 'UNKNOWN',
}
