// Configuration token for Microsoft Outlook OAuth
export const MICROSOFT_CONFIG = 'MICROSOFT_CONFIG';

// DI tokens for the pluggable shared-state backend (Redis or in-memory)
export const OUTLOOK_LOCK_STORE = 'OUTLOOK_LOCK_STORE';
export const OUTLOOK_RATE_LIMIT_STORE = 'OUTLOOK_RATE_LIMIT_STORE';

// Microsoft Graph API error codes
export const GRAPH_ERROR_CODES = {
  // this is happening when the mailbox is not enabled for REST API, soft-deleted, on-premise, inactive, etc.
  MAILBOX_NOT_ENABLED: 'MailboxNotEnabledForRESTAPI',
} as const;
