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
 * Entity for mapping users within a tenant connection for app-only authentication.
 *
 * Links external user IDs (from the host application) to Microsoft user IDs
 * for making Graph API calls on behalf of specific users within a tenant.
 *
 * This is an alias-compatible version that uses TenantConnection naming.
 * See also: MicrosoftTenantUser for the underlying implementation.
 */
@Entity('tenant_users')
export class TenantUser {
  @PrimaryGeneratedColumn('increment')
  id!: number;

  /**
   * The tenant this user belongs to.
   */
  @ManyToOne(() => MicrosoftTenant)
  @JoinColumn({ name: 'tenant_id' })
  tenant!: MicrosoftTenant;

  /**
   * Foreign key to the tenant.
   */
  @Column({ name: 'tenant_id' })
  tenantId!: number;

  /**
   * Microsoft Azure AD user object ID.
   * Used for /users/{id}/* Graph API calls.
   */
  @Column({ name: 'microsoft_user_id', length: 36 })
  @Index()
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
   * Max length 320 per RFC 5321 (64 local + @ + 255 domain).
   */
  @Column({ name: 'user_principal_name', length: 320 })
  userPrincipalName: string = '';

  /**
   * Default calendar ID for the user.
   * Cached to avoid repeated lookups.
   */
  @Column({ name: 'default_calendar_id', length: 255, nullable: true })
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
