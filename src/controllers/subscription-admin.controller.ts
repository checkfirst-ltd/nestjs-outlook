import { Controller, Post, Delete, HttpCode, Logger, Param, NotFoundException } from '@nestjs/common';
import { ApiTags, ApiResponse, ApiOperation, ApiParam } from '@nestjs/swagger';
import { MicrosoftSubscriptionService } from '../services/subscription/microsoft-subscription.service';

/**
 * Admin endpoints for on-demand subscription lifecycle operations.
 *
 * These endpoints invoke methods otherwise driven by cron / Graph lifecycle events:
 *   - POST   /subscriptions/admin/renew/:subscriptionId → renewWebhookSubscription
 *   - POST   /subscriptions/admin/health                → verifySubscriptionHealth (same as 6-hourly cron)
 *   - DELETE /subscriptions/admin/:subscriptionId/:userId → deleteWebhookSubscription
 */
@ApiTags('SubscriptionAdmin')
@Controller('subscriptions/admin')
export class SubscriptionAdminController {
  private readonly logger = new Logger(SubscriptionAdminController.name);

  constructor(
    private readonly subscriptionService: MicrosoftSubscriptionService,
  ) {}

  @Post('renew/:subscriptionId')
  @HttpCode(200)
  @ApiOperation({ summary: 'Renew a single subscription by ID.' })
  @ApiParam({ name: 'subscriptionId', description: 'The Microsoft Graph subscription ID to renew.' })
  @ApiResponse({ status: 200, description: 'Subscription renewed. Returns the new expiration timestamp.' })
  @ApiResponse({ status: 404, description: 'Subscription not found in local database.' })
  async triggerRenewOne(
    @Param('subscriptionId') subscriptionId: string,
  ): Promise<{ subscriptionId: string; newExpiration: string | null }> {
    this.logger.log(`[admin] Manual renewWebhookSubscription trigger for ${subscriptionId}`);
    const subscription = await this.subscriptionService.getSubscription(subscriptionId);
    if (!subscription) {
      throw new NotFoundException(`Subscription ${subscriptionId} not found in local database`);
    }
    const renewed = await this.subscriptionService.renewWebhookSubscription(
      subscriptionId,
      subscription.userId,
    );
    return {
      subscriptionId,
      newExpiration: renewed.expirationDateTime ?? null,
    };
  }

  @Post('health')
  @HttpCode(200)
  @ApiOperation({ summary: 'Run the subscription health-check job on demand (same as the 6-hourly cron).' })
  @ApiResponse({ status: 200, description: 'Health-check job completed. See logs for per-subscription outcomes.' })
  async triggerHealthCheck(): Promise<{ triggered: boolean; at: string }> {
    const at = new Date().toISOString();
    this.logger.log(`[admin] Manual verifySubscriptionHealth trigger at ${at}`);
    await this.subscriptionService.verifySubscriptionHealth();
    return { triggered: true, at };
  }

  @Delete(':subscriptionId/:userId')
  @HttpCode(200)
  @ApiOperation({ summary: 'Delete a subscription at Microsoft and deactivate it locally.' })
  @ApiParam({ name: 'subscriptionId', description: 'The Microsoft Graph subscription ID to delete.' })
  @ApiParam({ name: 'userId', description: 'External user ID (string) or internal user ID (numeric) owning the subscription.' })
  @ApiResponse({ status: 200, description: 'Subscription deleted (or already gone at Microsoft).' })
  async triggerDelete(
    @Param('subscriptionId') subscriptionId: string,
    @Param('userId') userId: string,
  ): Promise<{ subscriptionId: string; deleted: boolean }> {
    this.logger.log(`[admin] Manual deleteWebhookSubscription trigger for ${subscriptionId} (user ${userId})`);
    const numericUserId = Number(userId);
    const resolvedUserId = Number.isFinite(numericUserId) && String(numericUserId) === userId ? numericUserId : userId;
    const deleted = await this.subscriptionService.deleteWebhookSubscription(subscriptionId, resolvedUserId);
    return { subscriptionId, deleted };
  }
}
