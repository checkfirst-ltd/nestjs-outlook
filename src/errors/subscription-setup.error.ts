export enum SubscriptionFailureReason {
  PERMISSION_DENIED = 'PERMISSION_DENIED',
  AUTH_EXPIRED = 'AUTH_EXPIRED',
  RATE_LIMITED = 'RATE_LIMITED',
  NOT_FOUND = 'NOT_FOUND',
  SERVICE_UNAVAILABLE = 'SERVICE_UNAVAILABLE',
  UNKNOWN = 'UNKNOWN',
}

export class SubscriptionSetupError extends Error {
  constructor(
    message: string,
    public readonly reason: SubscriptionFailureReason = SubscriptionFailureReason.UNKNOWN,
  ) {
    super(message);
    this.name = 'SubscriptionSetupError';
  }
}
