import { ApiProperty, ApiPropertyOptional } from '@nestjs/swagger';
import { IsNotEmpty, IsString, IsDateString, IsOptional } from 'class-validator';

export class CreateTenantEventDto {
  @ApiProperty({
    description: 'External user ID from your application',
    example: 'user-123',
  })
  @IsString()
  @IsNotEmpty()
  externalUserId!: string;

  @ApiProperty({
    description: 'Event subject/title',
    example: 'Team Meeting',
  })
  @IsString()
  @IsNotEmpty()
  subject!: string;

  @ApiProperty({
    description: 'Event start date and time (ISO 8601 format)',
    example: '2025-01-15T10:00:00Z',
  })
  @IsDateString()
  @IsNotEmpty()
  startDateTime!: string;

  @ApiProperty({
    description: 'Event end date and time (ISO 8601 format)',
    example: '2025-01-15T11:00:00Z',
  })
  @IsDateString()
  @IsNotEmpty()
  endDateTime!: string;

  @ApiPropertyOptional({
    description: 'Event body/description (HTML supported)',
    example: 'Discussing Q1 goals',
  })
  @IsString()
  @IsOptional()
  body?: string;

  @ApiPropertyOptional({
    description: 'Event location',
    example: 'Conference Room A',
  })
  @IsString()
  @IsOptional()
  location?: string;
}
