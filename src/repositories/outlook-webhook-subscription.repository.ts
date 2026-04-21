import { Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { Repository, MoreThan } from 'typeorm';
import { OutlookWebhookSubscription } from '../entities/outlook-webhook-subscription.entity';
import { TtlCache } from '../utils/ttl-cache.util';

@Injectable()
export class OutlookWebhookSubscriptionRepository {
  private readonly bySubscriptionId = new TtlCache<string, OutlookWebhookSubscription>(60000);
  private readonly byUserId = new TtlCache<number, OutlookWebhookSubscription>(60000);

  constructor(
    @InjectRepository(OutlookWebhookSubscription)
    private readonly repository: Repository<OutlookWebhookSubscription>,
  ) {}

  private invalidate(subscription?: Partial<OutlookWebhookSubscription> | null): void {
    if (subscription?.subscriptionId) this.bySubscriptionId.delete(subscription.subscriptionId);
    if (subscription?.userId !== undefined) this.byUserId.delete(subscription.userId);
  }

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
        const saved = await this.repository.save(existingSubscription);
        this.invalidate(saved);
        return saved;
      }
    }

    // Create a new subscription if none exists
    // Use TypeORM's create method which safely copies properties while excluding 'id'
    const subscriptionWithoutId = { ...subscription };
    delete subscriptionWithoutId.id;
    const newSubscription = this.repository.create(subscriptionWithoutId);

    const saved = await this.repository.save(newSubscription);
    this.invalidate(saved);
    return saved;
  }

  async findBySubscriptionId(subscriptionId: string): Promise<OutlookWebhookSubscription | null> {
    const cached = this.bySubscriptionId.get(subscriptionId);
    if (cached !== undefined) return cached;

    const result = await this.repository.findOne({ where: { subscriptionId, isActive: true } });
    if (result) this.bySubscriptionId.set(subscriptionId, result);
    return result;
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
    this.bySubscriptionId.delete(subscriptionId);
    // userId unknown here; clear conservatively
    this.byUserId.clear();
  }

  async deactivateSubscription(subscriptionId: string): Promise<void> {
    await this.repository.update({ subscriptionId }, { isActive: false, updatedAt: new Date() });
    this.bySubscriptionId.delete(subscriptionId);
    this.byUserId.clear();
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
    const cached = this.byUserId.get(userId);
    if (cached !== undefined) return cached;

    const result = await this.repository.findOne({
      where: { userId, isActive: true },
    });
    if (result) this.byUserId.set(userId, result);
    return result;
  }

  async findAllActiveByUserIdAndResource(
    userId: number,
    resource: string,
  ): Promise<OutlookWebhookSubscription[]> {
    return this.repository.find({
      where: { userId, resource, isActive: true },
    });
  }

  async updateLastNotificationAt(subscriptionId: string): Promise<void> {
    await this.repository.update(
      { subscriptionId, isActive: true },
      { lastNotificationAt: new Date() },
    );
    this.bySubscriptionId.delete(subscriptionId);
  }

  async findStaleSubscriptions(staleThresholdHours: number): Promise<OutlookWebhookSubscription[]> {
    const threshold = new Date();
    threshold.setHours(threshold.getHours() - staleThresholdHours);

    return this.repository
      .createQueryBuilder('sub')
      .where('sub.isActive = :active', { active: true })
      .andWhere('sub.expirationDateTime > :now', { now: new Date() })
      .andWhere(
        '(sub.lastNotificationAt IS NULL OR sub.lastNotificationAt < :threshold)',
        { threshold },
      )
      .getMany();
  }

  async count(options: Parameters<Repository<OutlookWebhookSubscription>['count']>[0]): Promise<number> {
    return await this.repository.count(options);
  }
}
