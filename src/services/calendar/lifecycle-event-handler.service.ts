import { Injectable, Logger, Inject, forwardRef } from '@nestjs/common';
import { EventEmitter2 } from '@nestjs/event-emitter';
import { LifecycleEventType } from '../../enums/lifecycle-event-types.enum';
import { OutlookEventTypes } from '../../enums/event-types.enum';
import { OutlookWebhookSubscriptionRepository } from '../../repositories/outlook-webhook-subscription.repository';
import { CalendarService } from './calendar.service';
import { OutlookWebhookNotificationItemDto } from '../../dto/outlook-webhook-notification.dto';

/**
 * Service for handling Microsoft Graph lifecycle events
 *
 * Lifecycle events are sent by Microsoft Graph to notify about subscription
 * lifecycle changes and help reduce missing subscriptions and change notifications.
 *
 * @see https://learn.microsoft.com/en-us/graph/change-notifications-lifecycle-events
 */
@Injectable()
export class LifecycleEventHandlerService {
  private readonly logger = new Logger(LifecycleEventHandlerService.name);

  constructor(
    @Inject(forwardRef(() => CalendarService))
    private readonly calendarService: CalendarService,
    private readonly subscriptionRepository: OutlookWebhookSubscriptionRepository,
    private readonly eventEmitter: EventEmitter2,
  ) {}

  /**
   * Handle a lifecycle event notification
   *
   * @param notificationItem - The lifecycle event notification from Microsoft Graph
   */
  async handleLifecycleEvent(
    notificationItem: OutlookWebhookNotificationItemDto,
  ): Promise<{ success: boolean; message: string }> {
    const { lifecycleEvent, subscriptionId, tenantId } = notificationItem;

    if (!lifecycleEvent) {
      return {
        success: false,
        message: 'Missing lifecycleEvent field',
      };
    }

    if (!subscriptionId) {
      this.logger.warn('Lifecycle event received without subscriptionId');
      return {
        success: false,
        message: 'Missing subscriptionId field',
      };
    }

    this.logger.log(
      `Handling lifecycle event: ${lifecycleEvent} for subscription ${subscriptionId}`,
    );

    try {
      switch (lifecycleEvent) {
        case LifecycleEventType.REAUTHORIZATION_REQUIRED:
          return await this.handleReauthorizationRequired(
            subscriptionId,
            tenantId,
          );

        case LifecycleEventType.SUBSCRIPTION_REMOVED:
          return await this.handleSubscriptionRemoved(subscriptionId, tenantId);

        case LifecycleEventType.MISSED:
          return await this.handleMissedNotifications(subscriptionId, tenantId);

        default:
          this.logger.warn(`Unknown lifecycle event type: ${lifecycleEvent}`);
          return {
            success: false,
            message: `Unknown lifecycle event type: ${lifecycleEvent}`,
          };
      }
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(
        `Error handling lifecycle event ${lifecycleEvent}: ${errorMessage}`,
      );
      return {
        success: false,
        message: errorMessage,
      };
    }
  }

  /**
   * Handle reauthorizationRequired event
   *
   * This event indicates that:
   * - Access token is nearing expiration
   * - Subscription is approaching expiration
   * - Administrator revoked app permissions
   *
   * Action: Attempt to renew the subscription immediately
   */
  private async handleReauthorizationRequired(
    subscriptionId: string,
    tenantId?: string,
  ): Promise<{ success: boolean; message: string }> {
    this.logger.log(
      `[REAUTHORIZATION_REQUIRED] Subscription ${subscriptionId} requires reauthorization`,
    );

    try {
      // Find the subscription in our database
      const subscription =
        await this.subscriptionRepository.findBySubscriptionId(subscriptionId);

      if (!subscription) {
        this.logger.warn(
          `[REAUTHORIZATION_REQUIRED] Subscription ${subscriptionId} not found in database`,
        );
        return {
          success: false,
          message: 'Subscription not found',
        };
      }

      // Attempt to renew the subscription
      this.logger.log(
        `[REAUTHORIZATION_REQUIRED] Attempting to renew subscription ${subscriptionId}`,
      );

      const renewedSubscription =
        await this.calendarService.renewWebhookSubscriptionByUserId(
          subscriptionId,
          subscription.userId,
        );

      const expirationDate = renewedSubscription.expirationDateTime
        ? new Date(renewedSubscription.expirationDateTime).toISOString()
        : 'unknown';

      this.logger.log(
        `[REAUTHORIZATION_REQUIRED] Successfully renewed subscription ${subscriptionId}. New expiration: ${expirationDate}`,
      );

      // Emit event for application monitoring
      this.eventEmitter.emit(OutlookEventTypes.LIFECYCLE_REAUTHORIZATION_REQUIRED, {
        subscriptionId,
        tenantId,
        userId: subscription.userId,
        renewalSuccessful: true,
      });

      return {
        success: true,
        message: 'Subscription renewed successfully',
      };
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(
        `[REAUTHORIZATION_REQUIRED] Failed to renew subscription ${subscriptionId}: ${errorMessage}`,
      );

      // Emit event for application to notify user about manual reauthorization
      this.eventEmitter.emit(OutlookEventTypes.LIFECYCLE_REAUTHORIZATION_REQUIRED, {
        subscriptionId,
        tenantId,
        renewalSuccessful: false,
        error: errorMessage,
      });

      return {
        success: false,
        message: `Failed to renew subscription: ${errorMessage}`,
      };
    }
  }

  /**
   * Handle subscriptionRemoved event
   *
   * This event indicates that the subscription has been removed due to:
   * - Prolonged failure to deliver notifications
   * - Access token expiration without renewal
   *
   * Action: Mark subscription as inactive and clean up
   */
  private async handleSubscriptionRemoved(
    subscriptionId: string,
    tenantId?: string,
  ): Promise<{ success: boolean; message: string }> {
    this.logger.warn(
      `[SUBSCRIPTION_REMOVED] Subscription ${subscriptionId} has been removed by Microsoft Graph`,
    );

    try {
      // Find the subscription in our database
      const subscription =
        await this.subscriptionRepository.findBySubscriptionId(subscriptionId);

      if (!subscription) {
        this.logger.warn(
          `[SUBSCRIPTION_REMOVED] Subscription ${subscriptionId} not found in database`,
        );
        return {
          success: true,
          message: 'Subscription not found (already removed)',
        };
      }

      // Mark subscription as inactive
      await this.subscriptionRepository.deactivateSubscription(subscriptionId);

      this.logger.log(
        `[SUBSCRIPTION_REMOVED] Deactivated subscription ${subscriptionId} in database`,
      );

      // Emit event for application-level notification
      this.eventEmitter.emit(OutlookEventTypes.LIFECYCLE_SUBSCRIPTION_REMOVED, {
        subscriptionId,
        tenantId,
        userId: subscription.userId,
        resource: subscription.resource,
      });

      return {
        success: true,
        message: 'Subscription marked as inactive',
      };
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(
        `[SUBSCRIPTION_REMOVED] Error handling removed subscription ${subscriptionId}: ${errorMessage}`,
      );

      return {
        success: false,
        message: errorMessage,
      };
    }
  }

  /**
   * Handle missed notifications event
   *
   * This event indicates that change notifications were missed.
   * Action: Trigger a full delta sync to catch up on missed changes
   */
  private async handleMissedNotifications(
    subscriptionId: string,
    tenantId?: string,
  ): Promise<{ success: boolean; message: string }> {
    this.logger.warn(
      `[MISSED] Notifications were missed for subscription ${subscriptionId}. Triggering delta sync.`,
    );

    try {
      // Find the subscription in our database
      const subscription =
        await this.subscriptionRepository.findBySubscriptionId(subscriptionId);

      if (!subscription) {
        this.logger.warn(
          `[MISSED] Subscription ${subscriptionId} not found in database`,
        );
        return {
          success: false,
          message: 'Subscription not found',
        };
      }

      // Trigger delta sync to catch up on missed changes
      this.logger.log(
        `[MISSED] Starting delta sync for user ${subscription.userId} to recover missed changes`,
      );

      // Use streaming mode for better performance
      const changes = await this.calendarService.fetchAndSortChanges(
        String(subscription.userId),
        false, // Don't force reset, use existing delta link
      );

      const changesCount = changes.length;

      this.logger.log(
        `[MISSED] Delta sync completed. Found ${changesCount} changes for subscription ${subscriptionId}`,
      );

      // Emit event for monitoring
      this.eventEmitter.emit(OutlookEventTypes.LIFECYCLE_MISSED, {
        subscriptionId,
        tenantId,
        userId: subscription.userId,
        changesFound: changesCount,
      });

      return {
        success: true,
        message: `Delta sync completed. Found ${changesCount} changes.`,
      };
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(
        `[MISSED] Error during delta sync for subscription ${subscriptionId}: ${errorMessage}`,
      );

      // Emit event even on failure for monitoring
      this.eventEmitter.emit(OutlookEventTypes.LIFECYCLE_MISSED, {
        subscriptionId,
        tenantId,
        syncSuccessful: false,
        error: errorMessage,
      });

      return {
        success: false,
        message: `Delta sync failed: ${errorMessage}`,
      };
    }
  }
}
