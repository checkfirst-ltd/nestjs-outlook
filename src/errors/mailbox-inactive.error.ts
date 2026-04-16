export class MailboxInactiveError extends Error {
  constructor(public readonly graphMessage: string) {
    super(`Mailbox is not accessible: ${graphMessage}`);
    this.name = 'MailboxInactiveError';
  }
}
