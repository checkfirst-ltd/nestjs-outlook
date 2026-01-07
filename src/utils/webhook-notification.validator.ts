import { Logger } from '@nestjs/common';
import { OutlookWebhookNotificationItemDto } from '../dto/outlook-webhook-notification.dto';

/**
 * Resource types for webhook notifications
 */
export enum WebhookResourceType {
  CALENDAR = 'calendar',
  EMAIL = 'email',
}

/**
 * Validation result for webhook notifications
 */
export interface WebhookValidationResult {
  isValid: boolean;
  shouldSkip: boolean;
  isLifecycleEvent: boolean;
  lifecycleEventType?: string;
  reason?: string;
}

const EXPECTED_ODATA_TYPES = {
  [WebhookResourceType.CALENDAR]: '#microsoft.graph.event',
  [WebhookResourceType.EMAIL]: '#microsoft.graph.message',
} as const;

const LOG_PREFIXES = {
  [WebhookResourceType.CALENDAR]: '[CALENDAR_WEBHOOK]',
  [WebhookResourceType.EMAIL]: '[EMAIL_WEBHOOK]',
} as const;

/**
 * Validates a webhook notification item and provides consistent logging
 * for missing or invalid resourceData across calendar and email webhooks
 *
 * @param item - The notification item to validate
 * @param resourceType - The type of resource (calendar or email)
 * @param logger - Logger instance for outputting validation messages
 * @returns Validation result with isValid flag and optional reason
 */
export function validateNotificationItem(
  item: OutlookWebhookNotificationItemDto,
  resourceType: WebhookResourceType,
  logger: Logger,
): WebhookValidationResult {
  const logPrefix = LOG_PREFIXES[resourceType];
  const expectedODataType = EXPECTED_ODATA_TYPES[resourceType];

  // Check required fields
  if (!item.subscriptionId || !item.resource) {
    logger.warn(`${logPrefix} Missing required fields (subscriptionId or resource)`);
    return {
      isValid: false,
      shouldSkip: true,
      isLifecycleEvent: false,
      reason: 'Missing required fields',
    };
  }

  // Handle lifecycle events (these are special notifications that don't have changeType or resourceData)
  if (item.lifecycleEvent) {
    logger.log(
      `${logPrefix} Received lifecycle event: ${item.lifecycleEvent} for subscription ${item.subscriptionId}`
    );
    return {
      isValid: false,
      shouldSkip: true,
      isLifecycleEvent: true,
      lifecycleEventType: item.lifecycleEvent,
      reason: `Lifecycle event: ${item.lifecycleEvent}`,
    };
  }

  // Ensure changeType is present for non-lifecycle notifications
  if (!item.changeType) {
    logger.warn(
      `${logPrefix} Missing changeType for notification. ` +
      `Resource: ${item.resource}, SubscriptionId: ${item.subscriptionId}. ` +
      `This may indicate a malformed webhook payload.`
    );
    return {
      isValid: false,
      shouldSkip: true,
      isLifecycleEvent: false,
      reason: 'Missing changeType',
    };
  }

  const resourceData = item.resourceData;

  // Handle missing resourceData
  if (!resourceData) {
    if (item.changeType === 'deleted') {
      // For deleted resources, missing resourceData is potentially expected but concerning
      logger.warn(
        `${logPrefix} Missing resourceData for DELETED ${resourceType} - potential data loss. ` +
        `Resource: ${item.resource}, SubscriptionId: ${item.subscriptionId}, ChangeType: deleted`
      );
    } else {
      // For created/updated resources, this is critical - we're losing data
      logger.error(
        `${logPrefix} CRITICAL: Missing resourceData for ${item.changeType.toUpperCase()} ${resourceType} - DATA LOSS SCENARIO. ` +
        `Resource: ${item.resource}, SubscriptionId: ${item.subscriptionId}, ChangeType: ${item.changeType}. ` +
        `Manual sync recommended for this subscription.`
      );
    }
    return {
      isValid: false,
      shouldSkip: true,
      isLifecycleEvent: false,
      reason: 'Missing resourceData',
    };
  }

  // Validate @odata.type if present
  if (resourceData['@odata.type'] && resourceData['@odata.type'] !== expectedODataType) {
    logger.warn(
      `${logPrefix} Invalid resource type: ${resourceData['@odata.type']}. ` +
      `Expected: ${expectedODataType}, Resource: ${item.resource}`
    );
    return {
      isValid: false,
      shouldSkip: true,
      isLifecycleEvent: false,
      reason: `Invalid @odata.type: expected ${expectedODataType}`,
    };
  }

  // Valid notification
  return {
    isValid: true,
    shouldSkip: false,
    isLifecycleEvent: false,
  };
}

/**
 * Validates change type is supported
 *
 * @param changeType - The change type to validate
 * @param logger - Logger instance
 * @param logPrefix - Prefix for log messages
 * @returns true if valid, false otherwise
 */
export function validateChangeType(
  changeType: string,
  logger: Logger,
  logPrefix: string,
): boolean {
  const validChangeTypes = ['created', 'updated', 'deleted'];
  if (!validChangeTypes.includes(changeType)) {
    logger.warn(`${logPrefix} Unsupported change type: ${changeType}`);
    return false;
  }
  return true;
}
