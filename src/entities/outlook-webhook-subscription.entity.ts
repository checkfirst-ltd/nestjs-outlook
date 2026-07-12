import {
  Entity,
  PrimaryGeneratedColumn,
  Column,
  CreateDateColumn,
  UpdateDateColumn,
  Index,
} from 'typeorm';

@Entity('outlook_webhook_subscriptions')
export class OutlookWebhookSubscription {
  @PrimaryGeneratedColumn('increment')
  id!: number;

  @Column({ name: 'subscription_id', length: 255, unique: true })
  subscriptionId: string = '';

  @Column({ name: 'user_id' })
  userId: number = 0;

  /**
   * Microsoft tenant ID for app-only subscriptions.
   * Null for delegated (user) subscriptions.
   * Format: GUID (e.g., "12345678-1234-1234-1234-123456789abc")
   */
  @Column({ name: 'tenant_id', type: 'varchar', length: 36, nullable: true })
  @Index()
  tenantId: string | null = null;

  /**
   * Microsoft user ID (immutable ID) for app-only subscriptions.
   * Used in the resource path: /users/{microsoftUserId}/events
   * Null for delegated subscriptions which use /me/events.
   */
  @Column({ name: 'microsoft_user_id', type: 'varchar', length: 255, nullable: true })
  @Index()
  microsoftUserId: string | null = null;

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

  @Column({ name: 'last_notification_at', type: 'datetime', nullable: true })
  lastNotificationAt: Date | null = null;

  @CreateDateColumn({ name: 'created_at' })
  createdAt: Date = new Date();

  @UpdateDateColumn({ name: 'updated_at' })
  updatedAt: Date = new Date();
}
