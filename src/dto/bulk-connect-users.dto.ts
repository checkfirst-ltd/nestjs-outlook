import { IsString, IsOptional, IsArray, ArrayNotEmpty } from 'class-validator';
import { ApiProperty, ApiPropertyOptional } from '@nestjs/swagger';

/**
 * One user to bulk-connect into a tenant. `email` (or UPN) is required — the module resolves
 * the Microsoft account by email/UPN, so an external id alone is not enough.
 */
export class BulkConnectUserDto {
  @ApiProperty({
    description: "The host application's user identifier.",
    example: 'insp-001',
  })
  @IsString()
  externalUserId!: string;

  @ApiProperty({
    description: 'Email or user principal name (UPN) that exists in the tenant.',
    example: 'john.doe@contoso.com',
  })
  @IsString()
  email!: string;
}

/**
 * Request body for bulk-connecting users into a tenant.
 *
 * Note: nested-item validation is performed in the controller (this package does not depend on
 * `class-transformer`, so `@ValidateNested`/`@Type` are intentionally not used here).
 */
export class BulkConnectUsersDto {
  @ApiPropertyOptional({
    description: 'Azure AD tenant GUID. Defaults to the module-configured tenant when omitted.',
    example: '12345678-1234-1234-1234-123456789abc',
  })
  @IsOptional()
  @IsString()
  tenantId?: string;

  @ApiProperty({
    description: 'Users to connect into the tenant.',
    type: [BulkConnectUserDto],
  })
  @IsArray()
  @ArrayNotEmpty()
  users!: BulkConnectUserDto[];
}
