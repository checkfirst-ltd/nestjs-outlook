import {
  Entity,
  PrimaryGeneratedColumn,
  Column,
  CreateDateColumn,
  UpdateDateColumn,
  Index,
  ManyToOne,
  JoinColumn,
} from 'typeorm';
import { MicrosoftTenant } from './microsoft-tenant.entity';

/**
 * Entity for mapping users within a Microsoft tenant for app-only authentication.
 *
 * Links external user IDs (from the host application) to Microsoft user IDs
 * for making Graph API calls on behalf of specific users within a tenant.
 */
@Entity('microsoft_tenant_users')
export class MicrosoftTenantUser {
  @PrimaryGeneratedColumn('increment')
  id!: number;

  /**
   * The tenant this user belongs to.
   */
  @ManyToOne(() => MicrosoftTenant)
  @JoinColumn({ name: 'tenant_id' })
  tenant!: MicrosoftTenant;

  /**
   * Microsoft Azure AD user object ID.
   * Used for /users/{id}/* Graph API calls.
   */
  @Column({ name: 'microsoft_user_id', length: 36 })
  microsoftUserId: string = '';

  /**
   * External user ID from the host application.
   * Used to map host app users to Microsoft users.
   */
  @Column({ name: 'external_user_id', length: 255 })
  @Index()
  externalUserId: string = '';

  /**
   * User principal name (UPN) in the tenant.
   * Format: user@tenant.onmicrosoft.com
   */
  @Column({ name: 'user_principal_name', length: 255 })
  userPrincipalName: string = '';

  /**
   * Default calendar ID for the user.
   * Cached to avoid repeated lookups.
   */
  @Column({ name: 'default_calendar_id', type: 'varchar', length: 255, nullable: true })
  defaultCalendarId: string | null = null;

  /**
   * Whether the user mapping is active.
   */
  @Column({ name: 'is_active', default: true })
  isActive: boolean = true;

  @CreateDateColumn({ name: 'created_at' })
  createdAt: Date = new Date();

  @UpdateDateColumn({ name: 'updated_at' })
  updatedAt: Date = new Date();
}
