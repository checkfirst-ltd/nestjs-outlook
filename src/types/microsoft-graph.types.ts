/**
 * Re-exports of Microsoft Graph types used throughout the application
 * This provides a single import point for all Microsoft Graph types,
 * which makes it easier to maintain and allows for future extensions
 */

// Re-export types from Microsoft Graph
export type {
  // Calendar related types
  Event,
  Calendar,
  ItemBody,
  DateTimeTimeZone,
  Attendee,
  EmailAddress,
  Location,

  // Email related types
  Message,

  // Webhook related types
  Subscription,
  ChangeNotification,
  ChangeType
} from '@microsoft/microsoft-graph-types';

/**
 * Extended Event type that includes transactionId property
 * The transactionId maps to the iCalUId property in Microsoft Graph
 */
export interface EventWithTransactionId extends Event {
  transactionId?: string;
} 