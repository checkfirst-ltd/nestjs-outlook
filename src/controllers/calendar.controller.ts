import { Controller, Post, HttpCode, Query, Res, Req, Injectable, Body, Logger, UseGuards } from '@nestjs/common';
import { ApiTags, ApiResponse, ApiBody, ApiQuery } from '@nestjs/swagger';
import { Response, Request } from 'express';
import { randomUUID } from 'crypto';
import { CalendarService } from '../services/calendar/calendar.service';
import { ChangeNotification, ChangeType } from '@microsoft/microsoft-graph-types';
import { OutlookWebhookNotificationDto } from '../dto/outlook-webhook-notification.dto';
import { validateNotificationItem, validateChangeType, WebhookResourceType } from '../utils/webhook-notification.validator';
import { LifecycleEventHandlerService } from '../services/calendar/lifecycle-event-handler.service';
import { WebhookClientStateGuard, RequestWithWebhookValidation } from '../guards/webhook-client-state.guard';

@ApiTags('Calendar')
@Controller('calendar')
@Injectable()
export class CalendarController {
  private readonly logger = new Logger(CalendarController.name);

  constructor(
    private readonly calendarService: CalendarService,
    private readonly lifecycleEventHandler: LifecycleEventHandlerService,
  ) {}

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
  @Post('webhook/notification')
  @UseGuards(WebhookClientStateGuard)
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
  async handleCalendarWebhookNotification(
    @Query('validationToken') validationToken: string,
    @Body() notificationBody: OutlookWebhookNotificationDto,
    @Req() req: Request,
    @Res() res: Response,
  ) {
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
      const webhookTraceId = randomUUID();
      this.logger.log(`[WEBHOOK_RECEIVED] webhookTraceId=${webhookTraceId}, endpoint=notification, notificationCount=${notificationBody.value.length}`);
      this.logger.debug(`Received webhook notification: ${JSON.stringify(notificationBody)}`);

      // Drop items rejected by WebhookClientStateGuard (clientState / subscription security check)
      const { authorized, rejectedCount } = this.filterAuthorizedItems(
        req as RequestWithWebhookValidation,
        Array.isArray(notificationBody.value) ? notificationBody.value : [],
      );
      if (rejectedCount > 0) {
        this.logger.warn(`[SECURITY] Skipping ${rejectedCount.toString()} rejected webhook item(s), webhookTraceId=${webhookTraceId}`);
      }

      // Every item failed the security check - acknowledge with 200 but process nothing
      if (authorized.length === 0) {
        res.json({
          success: true,
          message: rejectedCount > 0
            ? `Rejected ${rejectedCount.toString()} notification(s) failing security validation`
            : `Received webhook notification with unexpected format`,
        });
        return;
      }

      // Early response with 202 Accepted if we have multiple notifications
      // This follows Microsoft's best practice to avoid timing out on the response
      if (authorized.length > 2) {
        this.logger.log(`Received batch of ${authorized.length.toString()} notifications, responding with 202 Accepted`);
        res.status(202).json({
          success: true,
          message: 'Notifications accepted for processing',
        });

        for (const item of authorized) {
          await this.calendarService.handleOutlookWebhookV2({
            subscriptionId: item.subscriptionId,
            subscriptionExpirationDateTime: item.subscriptionExpirationDateTime,
            changeType: item.changeType as ChangeType,
            resource: item.resource,
            resourceData: item.resourceData,
            clientState: item.clientState,
            tenantId: item.tenantId,
          }, webhookTraceId);
        }
        return;
      }

      // For smaller batches, process synchronously
      for (const item of authorized) {
        await this.calendarService.handleOutlookWebhookV2({
          subscriptionId: item.subscriptionId,
          subscriptionExpirationDateTime: item.subscriptionExpirationDateTime,
          changeType: item.changeType as ChangeType,
          resource: item.resource,
          resourceData: item.resourceData,
          clientState: item.clientState,
          tenantId: item.tenantId,
        }, webhookTraceId);
      }
      res.json({
        success: true,
        message: `Processed ${authorized.length.toString()} notifications`,
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
  @UseGuards(WebhookClientStateGuard)
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
      const webhookTraceId = randomUUID();
      this.logger.log(`[WEBHOOK_RECEIVED] webhookTraceId=${webhookTraceId}, endpoint=webhook, notificationCount=${notificationBody.value.length}`);
      this.logger.debug(`Received webhook notification: ${JSON.stringify(notificationBody)}`);

      // Drop items rejected by WebhookClientStateGuard (clientState / subscription security check)
      const { authorized, rejectedCount } = this.filterAuthorizedItems(
        req as RequestWithWebhookValidation,
        Array.isArray(notificationBody.value) ? notificationBody.value : [],
      );
      if (rejectedCount > 0) {
        this.logger.warn(`[SECURITY] Skipping ${rejectedCount.toString()} rejected webhook item(s), webhookTraceId=${webhookTraceId}`);
      }

      // Every item failed the security check - acknowledge with 200 but process nothing
      if (authorized.length === 0) {
        res.json({
          success: true,
          message: rejectedCount > 0
            ? `Rejected ${rejectedCount.toString()} notification(s) failing security validation`
            : `Received webhook notification with unexpected format`,
        });
        return;
      }

      // Early response with 202 Accepted if we have multiple notifications
      // This follows Microsoft's best practice to avoid timing out on the response
      if (authorized.length > 2) {
        this.logger.log(`Received batch of ${authorized.length.toString()} notifications, responding with 202 Accepted`);
        res.status(202).json({
          success: true,
          message: 'Notifications accepted for processing',
        });

        // Process notifications asynchronously after sending the response
        this.processCalendarNotificationBatch(authorized, webhookTraceId).catch((error: unknown) => {
          const errorMessage = error instanceof Error ? error.message : String(error);
          this.logger.error(`Error processing notification batch: ${errorMessage}`);
        });
        return;
      }

      // For smaller batches, process synchronously
      const results = await this.processCalendarNotificationBatch(authorized, webhookTraceId);
      res.json({
        success: true,
        message: `Processed ${results.successCount.toString()} out of ${authorized.length.toString()} notifications`,
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
   * Filter out notification items rejected by {@link WebhookClientStateGuard}.
   *
   * The guard returns 200 (so Microsoft stops retrying) but marks invalid items on the request.
   * This drops those items so they are never processed, while authorized items proceed normally.
   *
   * @param req Request carrying the guard's `webhookValidation` verdict
   * @param items The notification items in their original order
   * @returns The authorized items and how many were rejected
   */
  private filterAuthorizedItems<T>(
    req: RequestWithWebhookValidation,
    items: T[],
  ): { authorized: T[]; rejectedCount: number } {
    const validation = req.webhookValidation;
    if (!validation || validation.valid || validation.invalidItems.length === 0) {
      return { authorized: items, rejectedCount: 0 };
    }
    const invalidIndexes = new Set(validation.invalidItems.map((i) => i.index));
    const authorized = items.filter((_, idx) => !invalidIndexes.has(idx));
    return { authorized, rejectedCount: items.length - authorized.length };
  }

  /**
   * Process a batch of calendar notifications asynchronously
   * @param items The notification items to process (already filtered for authorization)
   * @param webhookTraceId Correlation ID for tracing this webhook through downstream operations
   * @returns Results of the processing operation
   */
  private async processCalendarNotificationBatch(
    items: OutlookWebhookNotificationDto['value'],
    webhookTraceId: string,
  ): Promise<{ successCount: number; failureCount: number }> {
    // Track processing results
    let successCount = 0;
    let failureCount = 0;

    // Track processed event IDs to avoid duplicates
    const processedEvents = new Set<string>();

    // Process each notification in the batch
    for (const item of items) {
      // Validate the notification item
      const validation = validateNotificationItem(
        item,
        WebhookResourceType.CALENDAR,
        this.logger
      );

      // Handle lifecycle events separately
      if (validation.isLifecycleEvent) {
        try {
          const result = await this.lifecycleEventHandler.handleLifecycleEvent(item);
          if (result.success) {
            this.logger.log(`Successfully handled lifecycle event: ${validation.lifecycleEventType}`);
            successCount++;
          } else {
            this.logger.warn(`Failed to handle lifecycle event: ${result.message}`);
            failureCount++;
          }
        } catch (error: unknown) {
          const errorMessage = error instanceof Error ? error.message : String(error);
          this.logger.error(`Error handling lifecycle event: ${errorMessage}`);
          failureCount++;
        }
        continue;
      }

      if (!validation.isValid || validation.shouldSkip) {
        failureCount++;
        continue;
      }

      // eslint-disable-next-line @typescript-eslint/no-non-null-assertion -- Validated above, guaranteed to be non-null
      const resourceData = item.resourceData!;

      // Skip duplicate events in the same batch
      if (resourceData.id && processedEvents.has(resourceData.id)) {
        this.logger.debug(`Skipping duplicate event: ${resourceData.id}`);
        continue;
      }

      // Add to processed set if it has an ID
      if (resourceData.id) {
        processedEvents.add(resourceData.id);
      }

      // Validate change type
      if (!validateChangeType(item.changeType || 'unknown', this.logger, '[CALENDAR_WEBHOOK]')) {
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

        const result = await this.calendarService.handleOutlookWebhook(changeNotification, false, webhookTraceId);

        if (result.success) {
          this.logger.log(`Successfully processed ${item.changeType} event for resource ID: ${resourceData.id || 'unknown'}, webhookTraceId=${webhookTraceId}`);
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