import { Injectable, Logger, NotFoundException } from '@nestjs/common';
import { OnEvent } from '@nestjs/event-emitter';
import { 
  OutlookEventTypes, 
  OutlookResourceData,
  CalendarService as MicrosoftCalendarService 
} from '@checkfirst/nestjs-outlook';
import { UserCalendarRepository } from './repositories/user-calendar.repository';
import { InjectRepository } from '@nestjs/typeorm';
import { Repository } from 'typeorm';
import { User } from '../users/entities/user.entity';

@Injectable()
export class CalendarService {
  private readonly logger = new Logger(CalendarService.name);

  constructor(
    private readonly userCalendarRepository: UserCalendarRepository,
    private readonly microsoftCalendarService: MicrosoftCalendarService,
    @InjectRepository(User)
    private readonly userRepository: Repository<User>,
  ) {}

  async createCalendarEvent(
    userId: number,
    eventData: {
      name: string;
      startDateTime: string;
      endDateTime: string;
    },
  ) {
    // Get the user's calendar data from database
    const userCalendar = await this.userCalendarRepository.findActiveByUserId(userId);
    
    if (!userCalendar) {
      throw new NotFoundException('No active Outlook calendar found for this user. Please connect to Outlook first.');
    }

    // Create the event object in Microsoft Graph format
    const event = {
      subject: eventData.name,
      start: {
        dateTime: eventData.startDateTime,
        timeZone: 'UTC',
      },
      end: {
        dateTime: eventData.endDateTime,
        timeZone: 'UTC',
      },
    };

    try {
      // Create the event using the Microsoft service
      // The service now handles token management internally
      const result = await this.microsoftCalendarService.createEvent(
        event,
        userCalendar.externalUserId,
        userCalendar.calendarId
      );

      return result.event;
    } catch (error) {
      this.logger.error(
        `Failed to create calendar event for user ${userId}: ${
          error instanceof Error ? error.message : 'Unknown error'
        }`,
      );
      throw error;
    }
  }

  @OnEvent(OutlookEventTypes.USER_AUTHENTICATED)
  async handleUserAuthenticated(externalUserId: string, data: { externalUserId: string, scopes: string[] }) {
    try {
      this.logger.log(`User authenticated: ${externalUserId}`);

      // First ensure the user exists
      const userIdNum = Number(externalUserId);
      let user = await this.userRepository.findOne({ where: { id: userIdNum } });
      
      if (!user) {
        // Create a basic user if it doesn't exist
        user = await this.userRepository.save({
          id: userIdNum,
          email: 'placeholder@email.com', // You might want to get this from Microsoft Graph API
          createdAt: new Date(),
          updatedAt: new Date(),
        });
        this.logger.log(`Created new user with ID ${externalUserId}`);
      }

      // Get the default calendar ID
      const calendarId = await this.microsoftCalendarService.getDefaultCalendarId(externalUserId);
      this.logger.log(`Retrieved default calendar ID: ${calendarId} for user ${externalUserId}`);

      // Get the existing calendar for this user if any
      const existingCalendar = await this.userCalendarRepository.findActiveByUserId(userIdNum);

      if (existingCalendar) {
        // Update existing calendar
        existingCalendar.externalUserId = externalUserId;
        existingCalendar.calendarId = calendarId;
        await this.userCalendarRepository.save(existingCalendar);
        this.logger.log(`Updated calendar for user ${externalUserId}`);
      } else {
        // Create new calendar entry
        await this.userCalendarRepository.saveCalendarDetails(
          userIdNum,
          externalUserId,
          calendarId
        );
        this.logger.log(`Created new calendar entry for user ${externalUserId} with calendar ID: ${calendarId}`);
      }
    } catch (error) {
      this.logger.error(
        `Failed to handle user authentication for user ${externalUserId}: ${
          error instanceof Error ? error.message : 'Unknown error'
        }`,
      );
    }
  }

  @OnEvent(OutlookEventTypes.EVENT_CREATED)
  handleOutlookEventCreated(data: OutlookResourceData) {
    this.logger.log(`New Outlook event created with ID: ${data.id}`);
  }

  @OnEvent(OutlookEventTypes.EVENT_DELETED)
  handleOutlookEventDeleted(data: OutlookResourceData) {
    this.logger.log(`Outlook event deleted with ID: ${data.id}`);
  }

  @OnEvent(OutlookEventTypes.EVENT_UPDATED)
  handleOutlookEventUpdated(data: OutlookResourceData) {
    this.logger.log(`Outlook event updated with ID: ${data.id}`);
  }
} 