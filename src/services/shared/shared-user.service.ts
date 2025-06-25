import { Logger } from '@nestjs/common';
import { Repository } from 'typeorm';
import { MicrosoftUser } from '../../entities/microsoft-user.entity';

export async function getExternalUserIdFromUserId(
  userId: number,
  microsoftUserRepository: Repository<MicrosoftUser>,
  logger: Logger
): Promise<string | null> {
  try {
    const user: MicrosoftUser | null = await microsoftUserRepository.findOne({ where: { id: userId } });
    if (user && typeof user.externalUserId === 'string') {
      return user.externalUserId;
    }
    return null;
  } catch (error) {
    logger.error(`Error getting externalUserId for userId ${String(userId)}:`, error);
    return null;
  }
} 