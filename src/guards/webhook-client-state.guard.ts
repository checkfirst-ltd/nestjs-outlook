import {
  Injectable,
  CanActivate,
  ExecutionContext,
  Logger,
} from '@nestjs/common';
import { EventEmitter2 } from '@nestjs/event-emitter';
import { Request } from 'express';
import { OutlookWebhookSubscriptionRepository } from '../repositories/outlook-webhook-subscription.repository';
import { OutlookEventTypes } from '../enums/event-types.enum';

/** Reasons a webhook notification item can fail the clientState security check. */
export type WebhookRejectionReason =
  | 'missing_subscription_id'
  | 'unknown_subscription'
  | 'missing_stored_client_state'
  | 'client_state_mismatch';

/** Per-item validation outcome attached to the request by {@link WebhookClientStateGuard}. */
export interface WebhookValidationItem {
  index: number;
  valid: boolean;
  reason?: WebhookRejectionReason;
  subscriptionId?: string;
  userId?: number | null;
}

/** Aggregate validation result attached to the request by {@link WebhookClientStateGuard}. */
export interface WebhookValidationResult {
  valid: boolean;
  invalidItems: WebhookValidationItem[];
}

/** Express request augmented with the guard's verdict for controllers to honor. */
export type RequestWithWebhookValidation = Request & {
  webhookValidation?: WebhookValidationResult;
};

/** Payload emitted on {@link OutlookEventTypes.WEBHOOK_REJECTED}. */
export interface WebhookRejectedEvent {
  subscriptionId: string;
  userId: number | null;
  reason: WebhookRejectionReason;
  endpoint: string;
}

/**
 * Guard that validates clientState in Microsoft Graph webhook notifications.
 *
 * Security behavior:
 * - Validation requests (with validationToken query param) are allowed through
 * - For notifications, each item's clientState is validated against stored value
 * - A subscription with no stored clientState is unverifiable and fails closed
 *   (rejected, not skipped)
 * - Mismatched/unverifiable clientState returns 200 (to stop Microsoft retries) but
 *   the controller skips processing invalid items (it reads `request.webhookValidation`)
 * - Each rejected item is logged as a security event AND emitted on
 *   {@link OutlookEventTypes.WEBHOOK_REJECTED} for observability
 *
 * @see https://learn.microsoft.com/en-us/graph/change-notifications-delivery-webhooks
 */
@Injectable()
export class WebhookClientStateGuard implements CanActivate {
  private readonly logger = new Logger(WebhookClientStateGuard.name);

  constructor(
    private readonly webhookSubscriptionRepository: OutlookWebhookSubscriptionRepository,
    private readonly eventEmitter: EventEmitter2,
  ) {}

  async canActivate(context: ExecutionContext): Promise<boolean> {
    const request = context.switchToHttp().getRequest<RequestWithWebhookValidation>();

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

    const endpoint = request.path;

    // Validate each notification item's clientState
    const validationResults: WebhookValidationItem[] = await Promise.all(
      body.value.map(async (item, index): Promise<WebhookValidationItem> => {
        const { subscriptionId, clientState } = item;

        if (!subscriptionId) {
          this.logger.warn(
            `[SECURITY] Notification item ${index} missing subscriptionId`
          );
          return { index, valid: false, reason: 'missing_subscription_id', userId: null };
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
          return { index, valid: false, reason: 'unknown_subscription', subscriptionId, userId: null };
        }

        // A stored subscription with no clientState is unverifiable. Fail closed
        // rather than skip the check — an empty stored value must never accept a
        // notification (see persistence: clientState is always minted non-empty).
        if (!subscription.clientState) {
          this.logger.warn(
            `[SECURITY] Subscription ${subscriptionId} has no stored clientState; ` +
            `notification cannot be verified. Rejecting as unverifiable.`
          );
          return {
            index,
            valid: false,
            reason: 'missing_stored_client_state',
            subscriptionId,
            userId: subscription.userId,
          };
        }

        // Validate clientState
        if (clientState !== subscription.clientState) {
          this.logger.warn(
            `[SECURITY] ClientState mismatch for subscription ${subscriptionId}. ` +
            `Expected prefix: user_${subscription.userId}_*, received: ${clientState?.substring(0, 20) ?? 'null'}...`
          );
          return {
            index,
            valid: false,
            reason: 'client_state_mismatch',
            subscriptionId,
            userId: subscription.userId,
          };
        }

        return { index, valid: true, subscriptionId, userId: subscription.userId };
      })
    );

    // Check if any validation failed
    const invalidItems = validationResults.filter((r) => !r.valid);

    if (invalidItems.length > 0) {
      this.logger.warn(
        `[SECURITY] Rejecting webhook: ${invalidItems.length}/${body.value.length} items failed validation. ` +
        `Reasons: ${invalidItems.map((i) => `item[${i.index}]:${i.reason ?? 'unknown'}`).join(', ')}`
      );

      // Emit one observability event per rejected item
      for (const item of invalidItems) {
        const payload: WebhookRejectedEvent = {
          subscriptionId: item.subscriptionId ?? 'unknown',
          userId: item.userId ?? null,
          reason: item.reason ?? 'unknown_subscription',
          endpoint,
        };
        this.eventEmitter.emit(OutlookEventTypes.WEBHOOK_REJECTED, payload);
      }

      // Attach verdict to the request; the controller returns 200 but skips invalid items.
      // Returning 200 (rather than 403) stops Microsoft from retrying forged notifications.
      request.webhookValidation = { valid: false, invalidItems };
      return true;
    }

    // All items valid
    request.webhookValidation = { valid: true, invalidItems: [] };

    this.logger.debug(
      `ClientState validated for ${body.value.length} notification item(s)`
    );

    return true;
  }
}
