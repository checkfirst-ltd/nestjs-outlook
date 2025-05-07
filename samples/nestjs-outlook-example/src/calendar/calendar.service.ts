import { Injectable, Logger, NotFoundException } from '@nestjs/common';
import { OnEvent } from '@nestjs/event-emitter';
import { OutlookEventTypes, OutlookResourceData, TokenResponse, OutlookService } from '@checkfirst/nestjs-outlook';
import { UserCalendarRepository } from './repositories/user-calendar.repository';
import { InjectRepository } from '@nestjs/typeorm';
import { Repository } from 'typeorm';
import { User } from '../users/entities/user.entity';

@Injectable()
export class CalendarService {
  private readonly logger = new Logger(CalendarService.name);

  constructor(
    private readonly userCalendarRepository: UserCalendarRepository,
    private readonly outlookService: OutlookService,
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
      // Create the event using stored credentials
      const result = await this.outlookService.createEvent(
        event,
        userCalendar.accessToken,
        userCalendar.refreshToken,
        userCalendar.tokenExpiry.toISOString(),
        userId,
        userCalendar.calendarId,
      );

      // If tokens were refreshed during the operation, update them in the database
      if (result.tokensRefreshed && result.refreshedTokens) {
        const tokenExpiry = new Date(Date.now() + (result.refreshedTokens.expires_in || 3600) * 1000);
        await this.userCalendarRepository.updateTokens(
          userCalendar.id,
          result.refreshedTokens.access_token,
          result.refreshedTokens.refresh_token,
          tokenExpiry,
        );
        this.logger.log(`Updated tokens for user ${userId} after refresh during event creation`);
      }

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

  @OnEvent(OutlookEventTypes.AUTH_TOKENS_SAVE)
  async handleAuthTokensSave(userId: string, tokenData: TokenResponse) {
    try {
      this.logger.log(`Saving tokens for user ${userId}`);

      // First ensure the user exists
      const userIdNum = Number(userId);
      let user = await this.userRepository.findOne({ where: { id: userIdNum } });
      
      if (!user) {
        // Create a basic user if it doesn't exist
        user = await this.userRepository.save({
          id: userIdNum,
          email: 'placeholder@email.com', // You might want to get this from Microsoft Graph API
          createdAt: new Date(),
          updatedAt: new Date(),
        });
        this.logger.log(`Created new user with ID ${userId}`);
      }

      // Calculate token expiry
      const expiresIn = tokenData.expires_in || 3600; // Default to 1 hour if not provided
      const tokenExpiry = new Date(Date.now() + expiresIn * 1000);

      // Get the default calendar ID from Microsoft Graph API
      const calendarId = await this.outlookService.getDefaultCalendarId(tokenData.access_token);
      this.logger.log(`Retrieved default calendar ID: ${calendarId} for user ${userId}`);

      // Get the existing calendar for this user if any
      const existingCalendar = await this.userCalendarRepository.findActiveByUserId(userIdNum);

      if (existingCalendar) {
        // Update existing calendar tokens
        await this.userCalendarRepository.updateTokens(
          existingCalendar.id,
          tokenData.access_token,
          tokenData.refresh_token,
          tokenExpiry,
        );
        this.logger.log(`Updated tokens for existing calendar for user ${userId}`);
      } else {
        // Create new calendar entry with the actual calendar ID
        await this.userCalendarRepository.saveCalendarCredentials(
          userIdNum,
          calendarId,
          tokenData.access_token,
          tokenData.refresh_token,
          tokenExpiry,
        );
        this.logger.log(`Created new calendar entry for user ${userId} with calendar ID: ${calendarId}`);
      }
    } catch (error) {
      this.logger.error(
        `Failed to save/update tokens for user ${userId}: ${
          error instanceof Error ? error.message : 'Unknown error'
        }`,
      );
      throw new Error('Failed to save Outlook calendar tokens');
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