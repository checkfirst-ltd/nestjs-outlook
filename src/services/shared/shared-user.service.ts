import { Logger } from '@nestjs/common';
import { Repository } from 'typeorm';
import { MicrosoftUser } from '../../entities/microsoft-user.entity';
import { TtlCache } from '../../utils/ttl-cache.util';

const externalUserIdCache = new TtlCache<number, string>(300000);

export async function getExternalUserIdFromUserId(
  userId: number,
  microsoftUserRepository: Repository<MicrosoftUser>,
  logger: Logger
): Promise<string | null> {
  const cached = externalUserIdCache.get(userId);
  if (cached !== undefined) return cached;

  try {
    const user: MicrosoftUser | null = await microsoftUserRepository.findOne({ where: { id: userId } });
    if (user && typeof user.externalUserId === 'string') {
      externalUserIdCache.set(userId, user.externalUserId);
      return user.externalUserId;
    }
    return null;
  } catch (error) {
    logger.error(`Error getting externalUserId for userId ${String(userId)}:`, error);
    return null;
  }
}
