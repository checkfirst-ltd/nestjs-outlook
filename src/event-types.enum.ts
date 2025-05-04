/**
 * Enum representing Microsoft Outlook event types used for NestJS event emitter
 */
export enum OutlookEventTypes {
  AUTH_TOKENS_SAVE = 'microsoft.auth.tokens.save',
  AUTH_TOKENS_UPDATE = 'microsoft.auth.tokens.update',
  EVENT_DELETED = 'outlook.event.deleted',
  EVENT_CREATED = 'outlook.event.created',
  EVENT_UPDATED = 'outlook.event.updated',
}
