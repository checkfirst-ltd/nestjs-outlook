import { Injectable, Logger, NotFoundException } from '@nestjs/common';
import { EmailService as MicrosoftEmailService, OutlookEventTypes, OutlookResourceData } from '@checkfirst/nestjs-outlook';
import { UserCalendarRepository } from '../calendar/repositories/user-calendar.repository';
import { Message, EmailAddress, BodyType } from '@microsoft/microsoft-graph-types';
import { OnEvent } from '@nestjs/event-emitter';

interface EmailRecipient {
  emailAddress: EmailAddress;
}

interface EmailData {
  subject?: string;
  receivedDateTime?: string;
  from?: EmailRecipient;
  toRecipients?: EmailRecipient[];
  body?: {
    content?: string;
    contentType?: BodyType;
  };
}

@Injectable()
export class EmailService {
  private readonly logger = new Logger(EmailService.name);

  constructor(
    private readonly userCalendarRepository: UserCalendarRepository,
    private readonly microsoftEmailService: MicrosoftEmailService,
  ) {}

  /**
   * Handle new incoming emails
   * This method is automatically called when a new email notification is received
   * @param data The email notification data
   */
  @OnEvent(OutlookEventTypes.EMAIL_RECEIVED)
  handleNewEmail(data: OutlookResourceData): void {
    try {
      this.logger.log(`📩 New email received for user ${data.userId ?? 'unknown'}`);

      // Extract email data if available
      if (data.data) {
        const emailData = data.data as EmailData;
        
        // Get basic email information
        const subject = emailData.subject ?? 'No subject';
        const receivedDateTime = emailData.receivedDateTime ?? 'Unknown time';
        const sender = emailData.from?.emailAddress.address ?? 'Unknown sender';
        
        // Get recipients
        const toRecipients = emailData.toRecipients?.map(r => r.emailAddress.address).join(', ') ?? 'No recipients';
        
        // Get email body (could be HTML or text)
        const bodyContent = emailData.body?.content ?? 'No content';
        const bodyType = emailData.body?.contentType ?? 'Unknown';
        
        // Print email details
        this.logger.debug(`
📧 EMAIL DETAILS:
-----------------
From: ${sender}
To: ${toRecipients}
Subject: ${subject}
Received: ${receivedDateTime}
Content Type: ${bodyType}
-----------------
Body Preview:
${bodyType === 'html' 
  ? 'HTML CONTENT: ' + bodyContent.substring(0, 100) + '...' 
  : bodyContent.substring(0, 200) + '...'}
-----------------
`);
      } else {
        this.logger.warn('Email notification received but email content was not available');
      }
    } catch (error) {
      this.logger.error(`Error processing new email: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }
  
  /**
   * Handle email updates
   * @param data The email notification data
   */
  @OnEvent(OutlookEventTypes.EMAIL_UPDATED)
  handleEmailUpdate(data: OutlookResourceData): void {
    this.logger.log(`📝 Email updated for user ${data.userId}, email ID: ${data.id || 'unknown'}`);
  }
  
  /**
   * Handle email deletions
   * @param data The email notification data
   */
  @OnEvent(OutlookEventTypes.EMAIL_DELETED)
  handleEmailDeletion(data: OutlookResourceData): void {
    this.logger.log(`🗑️ Email deleted for user ${data.userId}, email ID: ${data.id || 'unknown'}`);
  }

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
      await this.microsoftEmailService.sendEmail(
        message,
        userCalendar.externalUserId
      );

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