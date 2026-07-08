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
import { MicrosoftUserStatus } from '../enums/microsoft-user-status.enum';
import { MicrosoftTenant } from './microsoft-tenant.entity';

/**
 * Entity for storing Microsoft user information.
 *
 * A single row represents one host-application user (`externalUserId`) and can carry
 * either — or both — of two authentication capabilities:
 *
 * - **Delegated (per-user OAuth):** `accessToken` / `refreshToken` / `tokenExpiry` /
 *   `scopes` are populated when the user completes the OAuth flow. These track the
 *   specific scopes used during initial authentication so the same scopes are used on
 *   refresh.
 * - **App-only (tenant-wide):** `tenant` / `microsoftUserId` / `userPrincipalName` are
 *   populated when the host maps this user into a Microsoft tenant for app-only access.
 *   No per-user tokens are stored for this mode (the token is acquired at the tenant
 *   level), which is why the token columns are nullable.
 *
 * Keeping both capabilities on one row (keyed by `externalUserId`) means shared user
 * features — default calendar, active flag, future per-user settings — live in one place
 * regardless of how the user authenticated.
 */
@Entity('microsoft_users')
export class MicrosoftUser {
  @PrimaryGeneratedColumn('increment')
  id!: number;

  // Unique: one row per host user. Tenant mappings upsert onto this row, so delegated
  // lookups by external_user_id stay unambiguous (see DR-008).
  @Column({ name: 'external_user_id' })
  @Index({ unique: true })
  externalUserId: string = '';

  /**
   * Delegated OAuth access token. Null for users that only have app-only (tenant) access.
   */
  @Column({ name: 'access_token', type: 'text', nullable: true })
  accessToken: string | null = null;

  /**
   * Delegated OAuth refresh token. Null for users that only have app-only (tenant) access.
   */
  @Column({ name: 'refresh_token', type: 'text', nullable: true })
  refreshToken: string | null = null;

  @Column({ name: 'token_expiry', type: 'datetime', nullable: true })
  tokenExpiry: Date | null = null;

  @Column({ name: 'scopes', type: 'text', nullable: true })
  scopes: string | null = null;

  @Column({ name: 'is_active', default: true })
  isActive: boolean = true;

  @Column({
    name: 'status',
    type: 'varchar',
    length: 32,
    default: MicrosoftUserStatus.ACTIVE,
  })
  status: MicrosoftUserStatus = MicrosoftUserStatus.ACTIVE;

  @Column({ name: 'default_calendar_id', type: 'varchar', length: 255, nullable: true })
  defaultCalendarId: string | null = null;

  /**
   * The Microsoft tenant this user belongs to for app-only access.
   * Null for delegated-only users that were never mapped into a tenant.
   */
  @ManyToOne(() => MicrosoftTenant, { nullable: true, onDelete: 'CASCADE' })
  @JoinColumn({ name: 'tenant_id' })
  tenant: MicrosoftTenant | null = null;

  /**
   * Microsoft Azure AD user object ID, used for `/users/{id}/*` Graph API calls in
   * app-only mode. Null until the user is mapped into a tenant.
   */
  @Column({ name: 'microsoft_user_id', type: 'varchar', length: 36, nullable: true })
  @Index()
  microsoftUserId: string | null = null;

  /**
   * User principal name (UPN) in the tenant, e.g. `user@tenant.onmicrosoft.com`.
   * Max length 320 per RFC 5321 (64 local + @ + 255 domain). Null for delegated-only users.
   */
  @Column({ name: 'user_principal_name', type: 'varchar', length: 320, nullable: true })
  userPrincipalName: string | null = null;

  @CreateDateColumn({ name: 'created_at' })
  createdAt: Date = new Date();

  @UpdateDateColumn({ name: 'updated_at' })
  updatedAt: Date = new Date();
}
