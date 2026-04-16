// Configuration token for Microsoft Outlook OAuth
export const MICROSOFT_CONFIG = 'MICROSOFT_CONFIG';

// Microsoft Graph API error codes
export const GRAPH_ERROR_CODES = {
  // this is happening when the mailbox is not enabled for REST API, soft-deleted, on-premise, inactive, etc.
  MAILBOX_NOT_ENABLED: 'MailboxNotEnabledForRESTAPI',
} as const;
