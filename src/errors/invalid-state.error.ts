/**
 * Thrown when the OAuth `state` parameter is missing, malformed, or truncated
 * and cannot be parsed into a valid state object.
 *
 * This is a client-side (bad input) condition — the callback controller maps it
 * to HTTP 400 rather than 500, so bot/scanner traffic replaying chopped URLs
 * stops paging as server errors.
 */
export class InvalidStateError extends Error {
  constructor(public readonly reason: string) {
    super(`Invalid OAuth state parameter: ${reason}`);
    this.name = 'InvalidStateError';
  }
}
