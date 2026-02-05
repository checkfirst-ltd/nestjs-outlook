import { Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { Repository } from 'typeorm';
import { OutlookDeltaLink } from '../entities/delta-link.entity';
import { ResourceType } from '../enums/resource-type.enum';

@Injectable()
export class OutlookDeltaLinkRepository {
  constructor(
    @InjectRepository(OutlookDeltaLink)
    private readonly repository: Repository<OutlookDeltaLink>,
  ) {}

  /**
   * Save or update a delta link for a Microsoft user
   * @param internalUserId - Internal database user ID (MicrosoftUser.id)
   * @param resourceType - The type of resource (e.g., CALENDAR)
   * @param deltaLink - The delta link from Microsoft Graph API
   */
  async saveDeltaLink(
    internalUserId: number,
    resourceType: ResourceType,
    deltaLink: string,
  ): Promise<OutlookDeltaLink> {
    // Try to find an existing delta link for this user and resource type
    let deltaLinkEntity = await this.repository.findOne({
      where: { userId: internalUserId, resourceType },
    });

    // Create a new one if it doesn't exist
    if (!deltaLinkEntity) {
      deltaLinkEntity = new OutlookDeltaLink();
      deltaLinkEntity.userId = internalUserId;
      deltaLinkEntity.resourceType = resourceType;
    }

    // Update the delta link
    deltaLinkEntity.deltaLink = deltaLink;

    return this.repository.save(deltaLinkEntity);
  }

  /**
   * Get the delta link for a Microsoft user and resource type
   * @param internalUserId - Internal database user ID (MicrosoftUser.id)
   * @param resourceType - The type of resource (e.g., CALENDAR)
   * @returns The delta link or null if not found
   */
  async getDeltaLink(
    internalUserId: number,
    resourceType: ResourceType,
  ): Promise<string | null> {
    const deltaLinkEntity = await this.repository.findOne({
      where: { userId: internalUserId, resourceType },
      cache: 30000,
    });

    return deltaLinkEntity?.deltaLink || null;
  }

  /**
   * Delete the delta link for a Microsoft user and resource type
   * @param internalUserId - Internal database user ID (MicrosoftUser.id)
   * @param resourceType - The type of resource (e.g., CALENDAR)
   */
  async deleteDeltaLink(
    internalUserId: number,
    resourceType: ResourceType,
  ): Promise<void> {
    await this.repository.delete({ userId: internalUserId, resourceType });
  }
}