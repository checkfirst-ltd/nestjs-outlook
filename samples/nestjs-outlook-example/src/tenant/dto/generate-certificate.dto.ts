import { ApiProperty } from '@nestjs/swagger';
import { IsBoolean, IsOptional, IsString } from 'class-validator';

/**
 * Request to generate a self-signed certificate for app-only authentication.
 */
export class GenerateCertificateDto {
  @ApiProperty({
    description:
      'Azure AD tenant (directory) ID — a GUID. Used for the certificate CN and filename. ' +
      'Ignored when `shared` is true.',
    example: '12345678-1234-1234-1234-123456789abc',
    required: false,
  })
  @IsString()
  @IsOptional()
  tenantId?: string;

  @ApiProperty({
    description:
      'Generate the one-time SHARED Checkfirst app certificate instead of a per-tenant one. ' +
      'Operator-only setup step.',
    example: false,
    required: false,
  })
  @IsBoolean()
  @IsOptional()
  shared?: boolean;
}
