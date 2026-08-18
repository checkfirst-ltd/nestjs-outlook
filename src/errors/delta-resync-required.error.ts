/**
 * Thrown by the delta sync stream when Microsoft Graph rejects the stored delta
 * token with `410 Gone` (sync state expired / resync required) AND the caller
 * opted out of the library's automatic full-delta recovery
 * (`throwOnResyncRequired: true`).
 *
 * It lets a consumer choose a different recovery strategy than replaying the
 * whole mailbox through the delta API — e.g. a bounded windowed calendarView
 * import — instead of the library silently doing a full delta resync.
 *
 * The expired token is intentionally NOT deleted before this is thrown: the
 * caller owns recovery, and leaving the (already server-rejected) token in place
 * means a failed recovery simply re-signals on the next sync and retries, while
 * a successful recovery overwrites it via `initializeDeltaSync`.
 */
export class DeltaResyncRequiredError extends Error {
  constructor(
    public readonly externalUserId: string,
    public readonly resourceType: string = 'CALENDAR',
  ) {
    super(
      `Delta sync state expired (410 Gone) for user ${externalUserId} on ${resourceType}; ` +
        `a full resync is required (automatic delta recovery was disabled by the caller).`,
    );
    this.name = 'DeltaResyncRequiredError';
  }
}
