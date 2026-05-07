export class SubscriptionSetupError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'SubscriptionSetupError';
  }
}
