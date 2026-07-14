import { IsOptional, IsArray, ArrayNotEmpty, IsBoolean } from 'class-validator';
import { ApiProperty, ApiPropertyOptional } from '@nestjs/swagger';

/**
 * Request body for bulk health check / recover.
 *
 * Note: the `externalUserIds` items are validated in the controller (this package does not depend
 * on `class-transformer`, so array-of-primitives deep validation isn't done by a ValidationPipe).
 */
export class HealthCheckDto {
  @ApiProperty({
    description: "Host application user identifiers to check.",
    type: [String],
    example: ['insp-001', 'insp-002'],
  })
  @IsArray()
  @ArrayNotEmpty()
  externalUserIds!: string[];

  @ApiPropertyOptional({
    description: 'Also verify each subscription still exists at Microsoft Graph (extra Graph calls).',
    example: false,
  })
  @IsOptional()
  @IsBoolean()
  verifyAtGraph?: boolean;
}
