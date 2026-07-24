import { ApiProperty } from '@nestjs/swagger';
import { IsIn, IsNotEmpty, IsOptional, IsString } from 'class-validator';

export type TenantRegistrationMode = 'shared' | 'dedicated';

/**
 * Registers a Microsoft 365 tenant for app-only (client credentials) access.
 *
 * A tenant must be registered BEFORE an administrator runs the admin-consent
 * flow: the consent callback only activates an existing `microsoft_tenants`
 * row (it never creates one). The row is matched by `tenantId`, so the value
 * you register here must be the Azure AD directory ID and must also be passed
 * as the `state` of the admin-consent URL.
 */
export class RegisterTenantDto {
  @ApiProperty({
    description:
      'Onboarding model. "shared" uses the Checkfirst-owned app + shared certificate ' +
      '(only tenantId required). "dedicated" uses the tenant\'s own app registration and a ' +
      'per-tenant certificate (clientId + certificate fields required).',
    enum: ['shared', 'dedicated'],
    default: 'shared',
    required: false,
  })
  @IsIn(['shared', 'dedicated'])
  @IsOptional()
  mode?: TenantRegistrationMode;

  @ApiProperty({
    description: 'Azure AD tenant (directory) ID — a GUID',
    example: '12345678-1234-1234-1234-123456789abc',
  })
  @IsString()
  @IsNotEmpty()
  tenantId!: string;

  @ApiProperty({
    description:
      'Application (client) ID from the Azure AD app registration. Required in "dedicated" mode; ' +
      'in "shared" mode it defaults to the Checkfirst app\'s configured client ID.',
    example: '87654321-4321-4321-4321-cba987654321',
    required: false,
  })
  @IsString()
  @IsOptional()
  clientId?: string;

  @ApiProperty({
    description:
      'SHA-256 thumbprint of the certificate used for client-assertion signing. ' +
      'Optional when the application authenticates with a client secret instead.',
    example: 'A1B2C3D4E5F6A1B2C3D4E5F6A1B2C3D4E5F6A1B2',
    required: false,
  })
  @IsString()
  @IsOptional()
  certificateThumbprint?: string;

  @ApiProperty({
    description: 'Path to the X.509 certificate PEM file (optional)',
    required: false,
  })
  @IsString()
  @IsOptional()
  certificatePath?: string;

  @ApiProperty({
    description: 'Path to the private key PEM file (optional)',
    required: false,
  })
  @IsString()
  @IsOptional()
  certificateKeyPath?: string;
}
