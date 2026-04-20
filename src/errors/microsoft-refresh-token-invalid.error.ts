export class MicrosoftRefreshTokenInvalidError extends Error {
  constructor(public readonly internalUserId: number) {
    super(`Microsoft refresh token is invalid or expired for user ${internalUserId}`);
    this.name = 'MicrosoftRefreshTokenInvalidError';
  }
}
