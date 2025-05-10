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
  
  // Email events
  EMAIL_RECEIVED = 'outlook.email.received',
  EMAIL_UPDATED = 'outlook.email.updated',
  EMAIL_DELETED = 'outlook.email.deleted',
}
