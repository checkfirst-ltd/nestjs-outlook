import {
  Entity,
  PrimaryGeneratedColumn,
  Column,
  CreateDateColumn,
  UpdateDateColumn,
  Index,
} from 'typeorm';
import { MicrosoftTenantStatus } from '../enums/microsoft-tenant-status.enum';

/**
 * Entity for storing Microsoft tenant information for app-only authentication.
 *
 * This enables tenant-wide access to Microsoft Graph API using the OAuth 2.0
 * client credentials flow with certificate-based authentication.
 */
@Entity('microsoft_tenants')
export class MicrosoftTenant {
  @PrimaryGeneratedColumn('increment')
  id!: number;

  /**
   * Azure AD tenant ID (directory ID).
   * Format: GUID (e.g., "12345678-1234-1234-1234-123456789abc")
   */
  @Column({ name: 'tenant_id', length: 36, unique: true })
  @Index()
  tenantId: string = '';

  /**
   * Application (client) ID from Azure AD app registration for app-only auth.
   */
  @Column({ name: 'client_id', length: 36 })
  clientId: string = '';

  /**
   * SHA-256 thumbprint of the certificate (x5t#S256).
   * Used in the JWT header for certificate identification.
   */
  @Column({ name: 'certificate_thumbprint', length: 64 })
  certificateThumbprint: string = '';

  /**
   * Path to the X.509 certificate PEM file.
   * Used for client assertion signing.
   */
  @Column({ name: 'certificate_path', type: 'varchar', length: 255, nullable: true })
  certificatePath: string | null = null;

  /**
   * Path to the private key PEM file.
   * Must correspond to the public certificate.
   */
  @Column({ name: 'certificate_key_path', type: 'varchar', length: 255, nullable: true })
  certificateKeyPath: string | null = null;

  /**
   * Current status of the tenant connection.
   */
  @Column({
    name: 'status',
    type: 'varchar',
    length: 32,
    default: MicrosoftTenantStatus.PENDING_CONSENT,
  })
  status: MicrosoftTenantStatus = MicrosoftTenantStatus.PENDING_CONSENT;

  /**
   * Timestamp when admin consent was granted.
   */
  @Column({ name: 'admin_consent_granted_at', type: 'datetime', nullable: true })
  adminConsentGrantedAt: Date | null = null;

  /**
   * Whether the tenant is active and should be used.
   */
  @Column({ name: 'is_active', default: true })
  isActive: boolean = true;

  @CreateDateColumn({ name: 'created_at' })
  createdAt: Date = new Date();

  @UpdateDateColumn({ name: 'updated_at' })
  updatedAt: Date = new Date();
}
