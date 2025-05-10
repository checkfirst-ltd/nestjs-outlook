import {
  Entity,
  PrimaryGeneratedColumn,
  Column,
  CreateDateColumn,
  UpdateDateColumn,
} from 'typeorm';

@Entity('outlook_webhook_subscriptions')
export class OutlookWebhookSubscription {
  @PrimaryGeneratedColumn('increment')
  id!: number;

  @Column({ name: 'subscription_id', length: 255, unique: true })
  subscriptionId: string = '';

  @Column({ name: 'user_id' })
  userId: number = 0;

  @Column({ length: 255 })
  resource: string = '';

  @Column({ name: 'change_type', length: 255 })
  changeType: string = '';

  @Column({ name: 'client_state', length: 255 })
  clientState: string = '';

  @Column({ name: 'notification_url', length: 255 })
  notificationUrl: string = '';

  @Column({ name: 'expiration_date_time', type: 'datetime' })
  expirationDateTime: Date = new Date();

  @Column({ name: 'is_active', default: true })
  isActive: boolean = true;

  @CreateDateColumn({ name: 'created_at' })
  createdAt: Date = new Date();

  @UpdateDateColumn({ name: 'updated_at' })
  updatedAt: Date = new Date();
}
