import {
  Injectable,
  CanActivate,
  ExecutionContext,
  Logger,
} from '@nestjs/common';
import { Request } from 'express';
import { OutlookWebhookSubscriptionRepository } from '../repositories/outlook-webhook-subscription.repository';

/**
 * Guard that validates clientState in Microsoft Graph webhook notifications.
 *
 * Security behavior:
 * - Validation requests (with validationToken query param) are allowed through
 * - For notifications, each item's clientState is validated against stored value
 * - Mismatched clientState returns 200 (to stop Microsoft retries) but rejects processing
 * - Unknown subscriptions are logged as security events
 *
 * @see https://learn.microsoft.com/en-us/graph/change-notifications-delivery-webhooks
 */
@Injectable()
export class WebhookClientStateGuard implements CanActivate {
  private readonly logger = new Logger(WebhookClientStateGuard.name);

  constructor(
    private readonly webhookSubscriptionRepository: OutlookWebhookSubscriptionRepository,
  ) {}

  async canActivate(context: ExecutionContext): Promise<boolean> {
    const request = context.switchToHttp().getRequest<Request>();

    // Allow validation requests through (Microsoft endpoint verification)
    const validationToken = request.query.validationToken;
    if (validationToken) {
      this.logger.debug('Allowing validation request through');
      return true;
    }

    // Get notification body
    const body = request.body as { value?: Array<{ subscriptionId?: string; clientState?: string }> } | undefined;

    // No body or empty value array - allow through (will be handled by controller)
    if (!body?.value || !Array.isArray(body.value) || body.value.length === 0) {
      this.logger.debug('No notification items to validate');
      return true;
    }

    // Validate each notification item's clientState
    const validationResults = await Promise.all(
      body.value.map(async (item, index) => {
        const { subscriptionId, clientState } = item;

        if (!subscriptionId) {
          this.logger.warn(
            `[SECURITY] Notification item ${index} missing subscriptionId`
          );
          return { index, valid: false, reason: 'missing_subscription_id' };
        }

        // Look up stored subscription
        const subscription = await this.webhookSubscriptionRepository.findBySubscriptionId(
          subscriptionId
        );

        if (!subscription) {
          this.logger.warn(
            `[SECURITY] Unknown subscription ID: ${subscriptionId}. ` +
            `Possible replay attack or stale subscription.`
          );
          return { index, valid: false, reason: 'unknown_subscription' };
        }

        // Validate clientState
        if (subscription.clientState && clientState !== subscription.clientState) {
          this.logger.warn(
            `[SECURITY] ClientState mismatch for subscription ${subscriptionId}. ` +
            `Expected prefix: user_${subscription.userId}_*, received: ${clientState?.substring(0, 20) ?? 'null'}...`
          );
          return { index, valid: false, reason: 'client_state_mismatch' };
        }

        return { index, valid: true };
      })
    );

    // Check if any validation failed
    const invalidItems = validationResults.filter((r) => !r.valid);

    if (invalidItems.length > 0) {
      this.logger.warn(
        `[SECURITY] Rejecting webhook: ${invalidItems.length}/${body.value.length} items failed validation. ` +
        `Reasons: ${invalidItems.map((i) => `item[${i.index}]:${i.reason}`).join(', ')}`
      );

      // Attach validation results to request for controller to handle
      (request as Request & { webhookValidation: { valid: boolean; invalidItems: typeof invalidItems } }).webhookValidation = {
        valid: false,
        invalidItems,
      };

      // Return true but mark as invalid - controller should return 200 but not process
      // This stops Microsoft from retrying invalid notifications
      return true;
    }

    // All items valid
    (request as Request & { webhookValidation: { valid: boolean } }).webhookValidation = {
      valid: true,
    };

    this.logger.debug(
      `ClientState validated for ${body.value.length} notification item(s)`
    );

    return true;
  }
}
