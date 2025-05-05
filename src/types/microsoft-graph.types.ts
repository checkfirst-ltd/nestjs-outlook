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
  
  // Webhook related types
  Subscription,
  ChangeNotification,
  ChangeType
} from '@microsoft/microsoft-graph-types';

// You can extend or modify types if needed
// export interface ExtendedEvent extends Event {
//   customProperty?: string;
// } 