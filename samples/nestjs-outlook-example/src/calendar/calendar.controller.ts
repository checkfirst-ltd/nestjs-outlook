import { Controller, Get, Query, Req, UseGuards } from '@nestjs/common';
import { MicrosoftAuthService, OutlookService } from '@checkfirst/nestjs-outlook';
import { CalendarService } from './calendar.service';

interface UserTokenPayload {
  id: number;
  // Add other user properties as needed
}

@Controller('')
export class CalendarController {
  constructor(
    private readonly microsoftAuthService: MicrosoftAuthService,
    private readonly outlookService: OutlookService,
    private readonly calendarService: CalendarService,
  ) {}

  @Get('auth/microsoft/login')
  async login(@Req() req: any) {
    // In a real application, you would have proper authentication
    // For this example, we're using a mock user ID
    const mockUserId = '1';
    return await this.microsoftAuthService.getLoginUrl(mockUserId);
  }

  @Get('calendar-events')
  async createEvent(
    @Query('name') name: string,
    @Query('start-datetime') startDateTime: string,
    @Query('end-datetime') endDateTime: string,
  ) {
    // For this example, we're using mock tokens
    // In a real application, these would be stored securely and retrieved based on the authenticated user
    const mockUserId = 1;

    const event = {
      subject: name,
      start: {
        dateTime: startDateTime,
        timeZone: 'UTC',
      },
      end: {
        dateTime: endDateTime,
        timeZone: 'UTC',
      },
    };

    return await this.calendarService.createCalendarEvent(
      mockUserId,
      {
        name: name,
        startDateTime: startDateTime,
        endDateTime: endDateTime,
      },
    );
  }
} 