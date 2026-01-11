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
    // Check if a subscription with this subscriptionId already exists
    if (subscription.subscriptionId) {
      const existingSubscription = await this.repository.findOne({
        where: { subscriptionId: subscription.subscriptionId }
      });

      if (existingSubscription) {
        // Update the existing subscription but preserve the auto-generated id
        const originalId = existingSubscription.id;
        Object.assign(existingSubscription, subscription);
        existingSubscription.id = originalId; // Ensure ID doesn't get overwritten
        return this.repository.save(existingSubscription);
      }
    }

    // Create a new subscription if none exists
    // Use TypeORM's create method which safely copies properties while excluding 'id'
    const subscriptionWithoutId = { ...subscription };
    delete subscriptionWithoutId.id;
    const newSubscription = this.repository.create(subscriptionWithoutId);
    
    return this.repository.save(newSubscription);
  }

  async findBySubscriptionId(subscriptionId: string): Promise<OutlookWebhookSubscription | null> {
    return this.repository.findOne({ where: { subscriptionId, isActive: true } });
  }

  async updateSubscriptionExpiration(
    subscriptionId: string,
    expirationDateTime: Date,
  ): Promise<void> {
    const update: Partial<OutlookWebhookSubscription> = {
      expirationDateTime,
      updatedAt: new Date(),
    };

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

  async findActiveByUserId(userId: number): Promise<OutlookWebhookSubscription | null> {
    return this.repository.findOne({
      where: { userId, isActive: true },
    });
  }
}
