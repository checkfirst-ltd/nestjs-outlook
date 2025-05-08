import { Controller, Post, Body } from '@nestjs/common';
import { MicrosoftAuthService, CalendarService as MsCalendarService } from '@checkfirst/nestjs-outlook';
import { CalendarService } from './calendar.service';
import { CreateEventDto } from './dto/create-event.dto';
import { ApiTags, ApiOperation, ApiResponse } from '@nestjs/swagger';

@ApiTags('Calendar')
@Controller('calendar')
export class CalendarController {
  constructor(
    private readonly microsoftAuthService: MicrosoftAuthService,
    private readonly msCalendarService: MsCalendarService,
    private readonly calendarService: CalendarService,
  ) {}

  @Post('events')
  @ApiOperation({ summary: 'Create a new calendar event' })
  @ApiResponse({ 
    status: 201, 
    description: 'The event has been successfully created',
  })
  @ApiResponse({ status: 400, description: 'Invalid input data' })
  @ApiResponse({ status: 401, description: 'Unauthorized' })
  async createEvent(@Body() eventData: CreateEventDto) {
    // For this example, we're using mock tokens
    // In a real application, these would be stored securely and retrieved based on the authenticated user
    const mockUserId = 1;

    return await this.calendarService.createCalendarEvent(
      mockUserId,
      {
        name: eventData.name,
        startDateTime: eventData.startDateTime,
        endDateTime: eventData.endDateTime,
      },
    );
  }
} 