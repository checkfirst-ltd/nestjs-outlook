/**
 * Enum representing Microsoft Outlook event types used for NestJS event emitter
 */
export enum OutlookEventTypes {
  // Authentication events
  USER_AUTHENTICATED = 'microsoft.auth.user.authenticated',
  USER_REFRESH_TOKEN_INVALID = 'microsoft.auth.user.refresh_token_invalid',

  // Calendar events
  EVENT_DELETED = 'outlook.event.deleted',
  EVENT_CREATED = 'outlook.event.created',
  EVENT_UPDATED = 'outlook.event.updated',

  EVENT_NOTIFICATION = 'outlook.event.notification',

  // Calendar import events
  IMPORT_COMPLETED = 'outlook.calendar.import.completed',

  // Email events
  EMAIL_RECEIVED = 'outlook.email.received',
  EMAIL_UPDATED = 'outlook.email.updated',
  EMAIL_DELETED = 'outlook.email.deleted',

  // Lifecycle events
  LIFECYCLE_REAUTHORIZATION_REQUIRED = 'outlook.lifecycle.reauthorization_required',
  LIFECYCLE_SUBSCRIPTION_REMOVED = 'outlook.lifecycle.subscription_removed',
  LIFECYCLE_MISSED = 'outlook.lifecycle.missed',

  // Subscription lifecycle events
  SUBSCRIPTION_RECREATED = 'outlook.subscription.recreated',
  SUBSCRIPTION_RECREATION_FAILED = 'outlook.subscription.recreation_failed',
  SUBSCRIPTION_AUTH_FAILED = 'outlook.subscription.auth_failed',

  // Error events
  SUBSCRIPTION_CREATION_FAILED = 'outlook.subscription.creation_failed',

  // Cron job observability events
  HEALTH_CHECK_COMPLETED = 'outlook.cron.health_check.completed',
  RETRY_FAILED_SUBSCRIPTIONS_COMPLETED = 'outlook.cron.retry_failed_subscriptions.completed',

  // Security events
  WEBHOOK_REJECTED = 'outlook.webhook.rejected',

  // Tenant provisioning events
  TENANT_USERS_BULK_CONNECT_COMPLETED = 'outlook.tenant.users.bulk_connect.completed',
  TENANT_USERS_BULK_CONNECT_FAILED = 'outlook.tenant.users.bulk_connect.failed',

  // Health / recovery events
  USER_HEALTH_RECOVERY_COMPLETED = 'outlook.user.health.recovery.completed',
  USER_HEALTH_RECOVERY_FAILED = 'outlook.user.health.recovery.failed',
}
