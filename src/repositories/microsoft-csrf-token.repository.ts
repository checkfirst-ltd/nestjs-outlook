import { Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { Repository, LessThan } from 'typeorm';
import { MicrosoftCsrfToken } from '../entities/csrf-token.entity';

@Injectable()
export class MicrosoftCsrfTokenRepository {
  constructor(
    @InjectRepository(MicrosoftCsrfToken)
    private readonly repository: Repository<MicrosoftCsrfToken>,
  ) {}

  /**
   * Save a new CSRF token
   * @param token CSRF token
   * @param userId User ID
   * @param expiresInMs Expiration time in milliseconds
   */
  async saveToken(
    token: string,
    userId: string | number,
    expiresInMs: number,
  ): Promise<MicrosoftCsrfToken> {
    const expires = new Date(Date.now() + expiresInMs);

    const csrfToken = new MicrosoftCsrfToken();
    csrfToken.token = token;
    csrfToken.userId = userId.toString();
    csrfToken.expires = expires;

    return this.repository.save(csrfToken);
  }

  /**
   * Find and validate a token
   * @param token CSRF token
   * @returns The CSRF token entity if valid, null otherwise
   */
  async findAndValidateToken(token: string): Promise<MicrosoftCsrfToken | null> {
    if (!token) {
      return null;
    }

    // Find the token
    const csrfToken = await this.repository.findOne({
      where: { token },
    });

    if (!csrfToken) {
      return null;
    }

    // Check if token has expired
    if (csrfToken.expires < new Date()) {
      // Delete expired token
      await this.repository.delete(csrfToken.id);
      return null;
    }

    // Delete token after validation (one-time use)
    await this.repository.delete(csrfToken.id);

    return csrfToken;
  }

  /**
   * Clean up expired tokens
   */
  async cleanupExpiredTokens(): Promise<void> {
    const now = new Date();
    await this.repository.delete({ expires: LessThan(now) });
  }
}
