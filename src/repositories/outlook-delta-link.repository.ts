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

  async saveDeltaLink(
    userId: number,
    resourceType: ResourceType,
    deltaLink: string,
  ): Promise<OutlookDeltaLink> {
    // Try to find an existing delta link for this user and resource type
    let deltaLinkEntity = await this.repository.findOne({
      where: { userId, resourceType },
    });

    // Create a new one if it doesn't exist
    if (!deltaLinkEntity) {
      deltaLinkEntity = new OutlookDeltaLink();
      deltaLinkEntity.userId = userId;
      deltaLinkEntity.resourceType = resourceType;
    }

    // Update the delta link
    deltaLinkEntity.deltaLink = deltaLink;
    
    return this.repository.save(deltaLinkEntity);
  }

  async getDeltaLink(
    userId: number,
    resourceType: ResourceType,
  ): Promise<string | null> {
    const deltaLinkEntity = await this.repository.findOne({
      where: { userId, resourceType },
    });

    return deltaLinkEntity?.deltaLink || null;
  }

  async deleteDeltaLink(
    userId: number,
    resourceType: ResourceType,
  ): Promise<void> {
    await this.repository.delete({ userId, resourceType });
  }
}