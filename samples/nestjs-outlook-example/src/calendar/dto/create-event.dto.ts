import { IsNotEmpty, IsString, IsISO8601 } from 'class-validator';
import { ApiProperty } from '@nestjs/swagger';

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
    description: 'End date and time of the event in ISO 8601 format',
    example: '2023-04-20T15:00:00Z',
  })
  @IsNotEmpty()
  @IsISO8601()
  endDateTime: string;
} 