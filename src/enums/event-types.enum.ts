/**
 * Enum representing Microsoft Outlook event types used for NestJS event emitter
 */
export enum OutlookEventTypes {
  // Authentication events
  USER_AUTHENTICATED = 'microsoft.auth.user.authenticated',

  // Calendar events
  EVENT_DELETED = 'outlook.event.deleted',
  EVENT_CREATED = 'outlook.event.created',
  EVENT_UPDATED = 'outlook.event.updated',

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
}
