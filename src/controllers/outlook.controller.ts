import { Controller, Post, HttpCode, Query, Res, Req, Injectable, Body } from '@nestjs/common';
import { ApiTags, ApiResponse, ApiBody } from '@nestjs/swagger';
import { Response, Request } from 'express';
import { OutlookService } from '../services/outlook.service';
import { ChangeNotification, ChangeType } from '@microsoft/microsoft-graph-types';
import { OutlookWebhookNotificationDto } from '../dto/outlook-webhook-notification.dto';

@ApiTags('Outlook')
@Controller('outlook')
@Injectable()
export class OutlookController {
  constructor(private readonly outlookService: OutlookService) {}

  /**
   * Webhook endpoint for Outlook calendar notifications
   * This endpoint receives notifications when events are changed in Outlook
   * It also handles Microsoft Graph validation requests
   */
  @Post('webhook')
  @HttpCode(200)
  @ApiResponse({
    status: 200,
    description: 'Webhook notification processed successfully',
  })
  @ApiBody({
    description: 'Microsoft Graph webhook notification payload',
    type: OutlookWebhookNotificationDto,
    required: false,
  })
  async handleOutlookWebhook(
    @Query('validationToken') validationToken: string,
    @Body() notificationBody: OutlookWebhookNotificationDto,
    @Req() req: Request,
    @Res() res: Response,
  ): Promise<void> {
    // If this is a validation request, respond with the decoded validation token
    if (validationToken) {
      console.log('Handling Microsoft Graph validation request with token:', validationToken);
      // According to Microsoft's docs, we need to return the decoded token as plain text
      const decodedToken = decodeURIComponent(validationToken);
      res.set('Content-Type', 'text/plain; charset=utf-8');
      res.send(decodedToken);
      return;
    }

    // If not a validation request, process the notification
    try {
      console.log('Received webhook notification:', JSON.stringify(notificationBody));

      // Process each notification in the batch
      if (Array.isArray(notificationBody.value)) {
        // Keep track of processed events to avoid duplicate processing
        const processedEvents = new Set<string>();
        let hasSuccessfullyProcessed = false;

        for (const item of notificationBody.value) {
          const resourceData = item.resourceData;

          // Only process each event once
          if (resourceData.id && processedEvents.has(resourceData.id)) {
            continue;
          }

          // Add to processed set
          if (resourceData.id) {
            processedEvents.add(resourceData.id);
          }

          // Handle any change type (deleted, created, updated)
          if (resourceData.id && ['deleted', 'created', 'updated'].includes(item.changeType)) {
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

              const result = await this.outlookService.handleOutlookWebhook(changeNotification);
              console.log(
                `Processed ${item.changeType} event for ${resourceData.id}: ${JSON.stringify(result)}`,
              );
              hasSuccessfullyProcessed = true;
            } catch (error) {
              console.error(`Error processing ${item.changeType} event ${resourceData.id}:`, error);
            }
          } else {
            console.log(`Skipping notification of type: ${item.changeType}`);
          }
        }

        // Respond with success status
        res.json({
          success: true,
          message: hasSuccessfullyProcessed
            ? `Successfully processed events`
            : `Received notifications, but no events were processed`,
        });
        return;
      }

      // If the notification doesn't have the expected format, just acknowledge receipt
      res.json({
        success: true,
        message: `Received webhook notification in unexpected format`,
      });
    } catch (error) {
      console.error('Error processing webhook notification:', error);
      res.status(500).json({
        success: false,
        message: 'Error processing webhook notification',
      });
    }
  }
}
