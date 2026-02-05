import { Injectable, Logger } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { Repository } from 'typeorm';
import { MicrosoftUser } from '../../entities/microsoft-user.entity';

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
  async toInternalUserId(userId: string | number): Promise<number> {
    return typeof userId === 'string' ? await this.externalToInternal(userId) : userId;
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
  async externalToInternal(externalUserId: string): Promise<number> {
    const user = await this.microsoftUserRepository.findOne({
      where: { externalUserId },
      cache: 300000,
    });

    if (!user) {
      this.logger.error(
        `No active Microsoft user found for external ID: ${externalUserId}`,
      );
      throw new Error(
        `No active Microsoft user found for external ID: ${externalUserId}`,
      );
    }

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
  async internalToExternal(internalUserId: number): Promise<string> {
    const user = await this.microsoftUserRepository.findOne({
      where: { id: internalUserId },
      cache: 300000,
    });

    if (!user) {
      this.logger.error(
        `No Microsoft user found for internal ID: ${internalUserId}`,
      );
      throw new Error(
        `No Microsoft user found for internal ID: ${internalUserId}`,
      );
    }

    return user.externalUserId;
  }
}