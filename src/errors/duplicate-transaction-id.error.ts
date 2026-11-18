/**
 * Microsoft Graph refused to create an event because its `transactionId` was already used, and
 * the event that owns it could not be adopted in its place.
 *
 * Graph names the owner in `@event.existingEventId`. Creating an event normally adopts that
 * event instead of failing, so this error only surfaces when the owner is gone or Graph did
 * not name it. Retrying the same create cannot succeed.
 */
export class DuplicateTransactionIdError extends Error {
  constructor(
    public readonly transactionId: string | undefined,
    public readonly existingEventId: string | null,
    reason: string,
  ) {
    super(
      `TransactionId ${transactionId ?? '(none)'} is already used by ` +
        `${existingEventId ? `event ${existingEventId}` : 'an event Graph did not name'}: ${reason}`,
    );
    this.name = 'DuplicateTransactionIdError';
  }
}
