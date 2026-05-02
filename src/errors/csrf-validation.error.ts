export class CsrfValidationError extends Error {
  constructor(public readonly reason: string) {
    super(`CSRF validation failed: ${reason}`);
    this.name = 'CsrfValidationError';
  }
}
