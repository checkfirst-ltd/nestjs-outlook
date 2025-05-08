import { Injectable, Logger, NotFoundException } from '@nestjs/common';
import { EmailService as MicrosoftEmailService } from '@checkfirst/nestjs-outlook';
import { UserCalendarRepository } from '../calendar/repositories/user-calendar.repository';
import { Message } from '@microsoft/microsoft-graph-types';

@Injectable()
export class EmailService {
  private readonly logger = new Logger(EmailService.name);

  constructor(
    private readonly userCalendarRepository: UserCalendarRepository,
    private readonly microsoftEmailService: MicrosoftEmailService,
  ) {}

  async sendEmail(
    userId: number,
    to: string,
    subject: string,
    body: string,
  ) {
    // Get the user's Microsoft account data from database
    const userCalendar = await this.userCalendarRepository.findActiveByUserId(userId);
    
    if (!userCalendar) {
      throw new NotFoundException('No active Microsoft account found for this user. Please connect to Microsoft first.');
    }

    // Create the email message using Microsoft Graph format
    const message: Partial<Message> = {
      subject,
      body: {
        contentType: 'html',
        content: body
      },
      toRecipients: [
        {
          emailAddress: {
            address: to
          }
        }
      ]
    };

    try {
      // Send the email using the EmailService from nestjs-outlook
      const result = await this.microsoftEmailService.sendEmail(
        message,
        userCalendar.accessToken,
        userCalendar.refreshToken,
        userCalendar.tokenExpiry.toISOString(),
        userId
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
        this.logger.log(`Updated tokens for user ${userId} after refresh during email sending`);
      }

      return {
        success: true,
        message: 'Email sent successfully'
      };
    } catch (error) {
      this.logger.error(
        `Failed to send email for user ${userId}: ${
          error instanceof Error ? error.message : 'Unknown error'
        }`,
      );
      throw error;
    }
  }
} 