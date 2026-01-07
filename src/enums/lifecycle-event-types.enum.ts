/**
 * Microsoft Graph Lifecycle Event Types
 *
 * These events are sent by Microsoft Graph to notify about subscription lifecycle changes
 * and help reduce missing subscriptions and change notifications.
 *
 * @see https://learn.microsoft.com/en-us/graph/change-notifications-lifecycle-events
 */
export enum LifecycleEventType {
  /**
   * Indicates that the subscription is about to expire or requires reauthorization.
   * This can be triggered by:
   * - Access token nearing expiration
   * - Subscription approaching expiration
   * - Administrator revoking app permissions
   */
  REAUTHORIZATION_REQUIRED = 'reauthorizationRequired',

  /**
   * Indicates that the subscription has been removed.
   * This typically happens when:
   * - There's a prolonged failure to deliver notifications
   * - The access token has expired and wasn't renewed
   *
   * Supported for: Outlook messages, events, contacts, Teams chat messages
   */
  SUBSCRIPTION_REMOVED = 'subscriptionRemoved',

  /**
   * Indicates that change notifications were missed.
   * This notification helps trigger a full resync to catch up on missed changes.
   *
   * Supported for: Outlook messages, events, contacts
   */
  MISSED = 'missed',
}
