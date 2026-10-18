import {
  IsNotEmpty,
  IsString,
  IsISO8601,
  IsOptional,
  IsBoolean,
  IsEnum,
  ValidateIf,
} from 'class-validator';
import { ApiProperty, ApiPropertyOptional } from '@nestjs/swagger';
import { ShowAsType } from '@checkfirst/nestjs-outlook';

export class CreateEventDto {
  @ApiProperty({
    description: 'Name of the calendar event',
    example: 'Team Meeting',
  })
  @IsNotEmpty()
  @IsString()
  name: string;

  @ApiProperty({
    description: 'Start date and time of the event in ISO 8601 format',
    example: '2023-04-20T14:00:00Z',
  })
  @IsNotEmpty()
  @IsISO8601()
  startDateTime: string;

  @ApiProperty({
    description:
      'End date and time of the event in ISO 8601 format. When isAllDay is true, optional — calendar date is the last included day.',
    example: '2023-04-20T15:00:00Z',
  })
  @ValidateIf((dto: CreateEventDto) => dto.isAllDay !== true)
  @IsNotEmpty()
  @IsISO8601()
  endDateTime?: string;

  @ApiPropertyOptional({
    description:
      'When true, startDateTime/endDateTime are interpreted as all-day calendar dates (date portion only)',
    default: false,
  })
  @IsOptional()
  @IsBoolean()
  isAllDay?: boolean;

  @ApiPropertyOptional({
    description:
      'Free/busy status in Outlook (Graph showAs). Independent of isAllDay — pass "free" when you want the Free control in Outlook.',
    enum: ShowAsType,
    example: ShowAsType.FREE,
  })
  @IsOptional()
  @IsEnum(ShowAsType)
  showAs?: ShowAsType;
}
