import {
  Entity,
  PrimaryGeneratedColumn,
  Column,
  CreateDateColumn,
  UpdateDateColumn,
  Index,
} from 'typeorm';

/**
 * Entity for storing Microsoft user information including OAuth tokens
 * This helps track the specific scopes used during initial authentication
 * to ensure the same scopes are used during token refresh.
 */
@Entity('microsoft_users')
export class MicrosoftUser {
  @PrimaryGeneratedColumn('increment')
  id!: number;

  @Column({ name: 'external_user_id' })
  @Index()
  externalUserId: string = '';

  @Column({ name: 'access_token', type: 'text' })
  accessToken: string = '';

  @Column({ name: 'refresh_token', type: 'text' })
  refreshToken: string = '';

  @Column({ name: 'token_expiry', type: 'datetime' })
  tokenExpiry: Date = new Date();

  @Column({ name: 'scopes', type: 'text' })
  scopes: string = '';

  @Column({ name: 'is_active', default: true })
  isActive: boolean = true;

  @Column({ name: 'default_calendar_id', type: 'varchar', length: 255, nullable: true })
  defaultCalendarId: string | null = null;

  @CreateDateColumn({ name: 'created_at' })
  createdAt: Date = new Date();

  @UpdateDateColumn({ name: 'updated_at' })
  updatedAt: Date = new Date();
} 