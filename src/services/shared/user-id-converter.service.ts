import { Injectable, Logger } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { Repository } from 'typeorm';
import { MicrosoftUser } from '../../entities/microsoft-user.entity';
import { TtlCache } from '../../utils/ttl-cache.util';

/**
 * Service for converting between external user IDs and internal database IDs
 *
 * Terminology:
 * - External User ID: The user ID from the host application (string)
 * - Internal User ID: The auto-generated primary key in MicrosoftUser table (number)
 */
@Injectable()
export class UserIdConverterService {
  private readonly logger = new Logger(UserIdConverterService.name);
  private readonly externalToInternalCache = new TtlCache<string, number>(5000);
  private readonly internalToExternalCache = new TtlCache<number, string>(3000);

  constructor(
    @InjectRepository(MicrosoftUser)
    private readonly microsoftUserRepository: Repository<MicrosoftUser>,
  ) {}

  /**
   * 
   * @param userId - The user ID from the host application or the internal database ID
   * @returns The internal database ID
   * @throws Error if no Microsoft user found for the external ID
   *
   * @example
   * ```typescript
   * // Your app user ID: "7"
   * const internalId = await converter.toInternalUserId("7");
   * // Returns: 42 (database primary key)
   * ```
   */
  async toInternalUserId(userId: string | number, {cache = true}: {cache?: boolean} = {}): Promise<number> {
    return typeof userId === 'string' ? await this.externalToInternal(userId, {cache}) : userId;
  }

  /**
   * Convert external user ID to internal database ID
   *
   * @param externalUserId - The user ID from the host application
   * @returns The internal database ID (primary key)
   * @throws Error if no active Microsoft user found for the external ID
   *
   * @example
   * ```typescript
   * // Your app user ID: "7"
   * const internalId = await converter.externalToInternal("7");
   * // Returns: 42 (database primary key)
   * ```
   */
  async externalToInternal(externalUserId: string, {cache = true}: {cache?: boolean} = {}): Promise<number> {
    if (cache) {
      const cached = this.externalToInternalCache.get(externalUserId);
      if (cached !== undefined) return cached;
    }

    const user = await this.microsoftUserRepository.findOne({
      where: { externalUserId },
    });

    if (!user) {
      this.logger.error(
        `No active Microsoft user found for external ID: ${externalUserId}`,
      );
      throw new Error(
        `No active Microsoft user found for external ID: ${externalUserId}`,
      );
    }

    if (cache) this.externalToInternalCache.set(externalUserId, user.id);
    return user.id;
  }

  /**
   * Convert internal database ID to external user ID
   *
   * @param internalUserId - The internal database ID (primary key)
   * @returns The external user ID from the host application
   * @throws Error if no Microsoft user found for the internal ID
   *
   * @example
   * ```typescript
   * // Database primary key: 42
   * const externalId = await converter.internalToExternal(42);
   * // Returns: "7" (your app's user ID)
   * ```
   */
  async internalToExternal(internalUserId: number, {cache = true}: {cache?: boolean} = {}): Promise<string> {
    if (cache) {
      const cached = this.internalToExternalCache.get(internalUserId);
      if (cached !== undefined) return cached;
    }

    const user = await this.microsoftUserRepository.findOne({
      where: { id: internalUserId },
    });

    if (!user) {
      this.logger.error(
        `No Microsoft user found for internal ID: ${internalUserId}`,
      );
      throw new Error(
        `No Microsoft user found for internal ID: ${internalUserId}`,
      );
    }

    if (cache) this.internalToExternalCache.set(internalUserId, user.externalUserId);
    return user.externalUserId;
  }
}