import { Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { Repository } from 'typeorm';
import { OutlookDeltaLink } from '../entities/delta-link.entity';

@Injectable()
export class OutlookDeltaLinkRepository {
  constructor(
    @InjectRepository(OutlookDeltaLink)
    private readonly repository: Repository<OutlookDeltaLink>,
  ) {}

  async saveDeltaLink(
    externalUserId: string,
    resourceType: string,
    deltaLink: string,
  ): Promise<OutlookDeltaLink> {
    // Try to find an existing delta link for this user and resource type
    let deltaLinkEntity = await this.repository.findOne({
      where: { externalUserId, resourceType },
    });

    // Create a new one if it doesn't exist
    if (!deltaLinkEntity) {
      deltaLinkEntity = new OutlookDeltaLink();
      deltaLinkEntity.externalUserId = externalUserId;
      deltaLinkEntity.resourceType = resourceType;
    }

    // Update the delta link
    deltaLinkEntity.deltaLink = deltaLink;
    
    return this.repository.save(deltaLinkEntity);
  }

  async getDeltaLink(
    externalUserId: string,
    resourceType: string,
  ): Promise<string | null> {
    const deltaLinkEntity = await this.repository.findOne({
      where: { externalUserId, resourceType },
    });

    return deltaLinkEntity?.deltaLink || null;
  }
}