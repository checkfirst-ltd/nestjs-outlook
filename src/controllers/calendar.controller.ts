import { Controller, Post, HttpCode, Query, Res, Req, Injectable, Body, Logger } from '@nestjs/common';
import { ApiTags, ApiResponse, ApiBody, ApiQuery } from '@nestjs/swagger';
import { Response, Request } from 'express';
import { CalendarService } from '../services/calendar/calendar.service';
import { ChangeNotification, ChangeType } from '@microsoft/microsoft-graph-types';
import { OutlookWebhookNotificationDto } from '../dto/outlook-webhook-notification.dto';

@ApiTags('Calendar')
@Controller('calendar')
@Injectable()
export class CalendarController {
  private readonly logger = new Logger(CalendarController.name);
  
  constructor(private readonly calendarService: CalendarService) {}

  /**
   * Webhook endpoint for Outlook calendar notifications
   * 
   * This endpoint receives notifications when calendar events are changed in Outlook
   * and handles Microsoft Graph validation requests.
   * 
   * It follows Microsoft's best practices for webhook implementations:
   * - Responds within 3 seconds or returns 202 Accepted for long-running processes
   * - Properly handles validation requests
   * - Returns appropriate HTTP status codes
   * 
   * @see https://learn.microsoft.com/en-us/graph/change-notifications-delivery-webhooks
   */
  @Post('webhook')
  @HttpCode(200)
  @ApiResponse({
    status: 200,
    description: 'Webhook validation or notification processed successfully',
  })
  @ApiResponse({
    status: 202,
    description: 'Notification accepted for processing',
  })
  @ApiResponse({
    status: 500,
    description: 'Server error processing the notification',
  })
  @ApiQuery({
    name: 'validationToken',
    required: false,
    description: 'Token sent by Microsoft Graph to validate the webhook endpoint',
  })
  @ApiBody({
    description: 'Microsoft Graph webhook notification payload',
    type: OutlookWebhookNotificationDto,
    required: false,
  })
  async handleCalendarWebhook(
    @Query('validationToken') validationToken: string,
    @Body() notificationBody: OutlookWebhookNotificationDto,
    @Req() req: Request,
    @Res() res: Response,
  ): Promise<void> {
    // Handle Microsoft Graph endpoint validation
    if (validationToken) {
      this.logger.log('Handling Microsoft Graph validation request');
      
      // According to Microsoft's docs, we need to return the decoded token as plain text
      const decodedToken = decodeURIComponent(validationToken);
      res.set('Content-Type', 'text/plain; charset=utf-8');
      res.send(decodedToken);
      return;
    }

    // Process notification
    try {
      this.logger.debug(`Received webhook notification: ${JSON.stringify(notificationBody)}`);
      
      // Early response with 202 Accepted if we have multiple notifications
      // This follows Microsoft's best practice to avoid timing out on the response
      if (Array.isArray(notificationBody.value) && notificationBody.value.length > 2) {
        this.logger.log(`Received batch of ${notificationBody.value.length.toString()} notifications, responding with 202 Accepted`);
        res.status(202).json({
          success: true,
          message: 'Notifications accepted for processing',
        });
        
        // Process notifications asynchronously after sending the response
        this.processCalendarNotificationBatch(notificationBody).catch((error: unknown) => {
          const errorMessage = error instanceof Error ? error.message : String(error);
          this.logger.error(`Error processing notification batch: ${errorMessage}`);
        });
        return;
      }
      
      // For smaller batches, process synchronously
      if (Array.isArray(notificationBody.value) && notificationBody.value.length > 0) {
        const results = await this.processCalendarNotificationBatch(notificationBody);
        res.json({
          success: true,
          message: `Processed ${results.successCount.toString()} out of ${notificationBody.value.length.toString()} notifications`,
        });
        return;
      }

      // Handle empty or unexpected notification format
      this.logger.warn('Received webhook notification with unexpected format');
      res.json({
        success: true,
        message: `Received webhook notification with unexpected format`,
      });
    } catch (error: unknown) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      const errorStack = error instanceof Error ? error.stack : undefined;
      this.logger.error(`Error processing webhook notification: ${errorMessage}`, errorStack);
      res.status(500).json({
        success: false,
        message: 'Error processing webhook notification',
      });
    }
  }
  
  /**
   * Process a batch of calendar notifications asynchronously
   * @param notificationBody The batch of notifications to process
   * @returns Results of the processing operation
   */
  private async processCalendarNotificationBatch(
    notificationBody: OutlookWebhookNotificationDto,
  ): Promise<{ successCount: number; failureCount: number }> {
    // Track processing results
    let successCount = 0;
    let failureCount = 0;
    
    // Track processed event IDs to avoid duplicates
    const processedEvents = new Set<string>();
    
    // Process each notification in the batch
    for (const item of notificationBody.value) {
      // Skip invalid notifications
      if (!item.subscriptionId || !item.resource) {
        this.logger.warn(`Skipping notification with missing required fields`);
        failureCount++;
        continue;
      }
      
      const resourceData = item.resourceData;

      // Skip notifications without resourceData
      if (!resourceData) {
        this.logger.warn(`Skipping notification with missing resourceData`);
        failureCount++;
        continue;
      }

      // Skip duplicate events in the same batch
      if (resourceData.id && processedEvents.has(resourceData.id)) {
        this.logger.debug(`Skipping duplicate event: ${resourceData.id}`);
        continue;
      }

      // Add to processed set if it has an ID
      if (resourceData.id) {
        processedEvents.add(resourceData.id);
      }

      // Only process notifications with supported change types
      const validChangeTypes = ['created', 'updated', 'deleted'];
      if (!validChangeTypes.includes(item.changeType)) {
        this.logger.warn(`Skipping notification with unsupported change type: ${item.changeType}`);
        failureCount++;
        continue;
      }
      
      try {
        // Convert to ChangeNotification type for the service
        const changeNotification: ChangeNotification = {
          subscriptionId: item.subscriptionId,
          subscriptionExpirationDateTime: item.subscriptionExpirationDateTime,
          changeType: item.changeType as ChangeType,
          resource: item.resource,
          resourceData: resourceData,
          clientState: item.clientState,
          tenantId: item.tenantId,
        };

        const result = await this.calendarService.handleOutlookWebhook(changeNotification);
        
        if (result.success) {
          this.logger.log(`Successfully processed ${item.changeType} event for resource ID: ${resourceData.id || 'unknown'}`);
          successCount++;
        } else {
          this.logger.warn(`Failed to process event: ${result.message}`);
          failureCount++;
        }
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        this.logger.error(`Error processing event: ${errorMessage}`);
        failureCount++;
      }
    }
    
    this.logger.log(`Finished processing batch: ${successCount.toString()} succeeded, ${failureCount.toString()} failed`);
    return { successCount, failureCount };
  }
} 