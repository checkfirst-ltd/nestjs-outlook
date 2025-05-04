import { Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { Repository, LessThan, MoreThan } from 'typeorm';
import { OutlookWebhookSubscription } from '../entities/outlook-webhook-subscription.entity';

@Injectable()
export class OutlookWebhookSubscriptionRepository {
  constructor(
    @InjectRepository(OutlookWebhookSubscription)
    private readonly repository: Repository<OutlookWebhookSubscription>,
  ) {}

  async saveSubscription(
    subscription: Partial<OutlookWebhookSubscription>,
  ): Promise<OutlookWebhookSubscription> {
    const newSubscription = this.repository.create(subscription);
    return this.repository.save(newSubscription);
  }

  async findBySubscriptionId(subscriptionId: string): Promise<OutlookWebhookSubscription | null> {
    return this.repository.findOne({ where: { subscriptionId, isActive: true } });
  }

  async updateSubscriptionExpiration(
    subscriptionId: string,
    expirationDateTime: Date,
    accessToken?: string,
  ): Promise<void> {
    const update: Partial<OutlookWebhookSubscription> = {
      expirationDateTime,
      updatedAt: new Date(),
    };

    if (accessToken) {
      update.accessToken = accessToken;
    }

    await this.repository.update({ subscriptionId, isActive: true }, update);
  }

  async deactivateSubscription(subscriptionId: string): Promise<void> {
    await this.repository.update({ subscriptionId }, { isActive: false, updatedAt: new Date() });
  }

  async findSubscriptionsNeedingRenewal(
    hoursUntilExpiration: number,
  ): Promise<OutlookWebhookSubscription[]> {
    const expirationThreshold = new Date();
    expirationThreshold.setHours(expirationThreshold.getHours() + hoursUntilExpiration);

    return this.repository.find({
      where: {
        isActive: true,
        expirationDateTime: LessThan(expirationThreshold),
      },
    });
  }

  async findActiveSubscriptions(): Promise<OutlookWebhookSubscription[]> {
    const now = new Date();
    return this.repository.find({
      where: {
        isActive: true,
        expirationDateTime: MoreThan(now),
      },
    });
  }
}
