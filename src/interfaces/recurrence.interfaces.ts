/**
 * Recurrence-related interfaces for Outlook recurring events.
 *
 * These types model Microsoft Graph's recurring event data and provide
 * a clean contract between nestjs-outlook (provider) and nestjs-calendar-hub (consumer).
 */

/**
 * Mirrors Microsoft Graph PatternedRecurrence.
 * Stored as JSON on series master rows.
 */
export interface RecurrenceRule {
  pattern: {
    type:
      | 'daily'
      | 'weekly'
      | 'absoluteMonthly'
      | 'relativeMonthly'
      | 'absoluteYearly'
      | 'relativeYearly';
    interval: number;
    daysOfWeek?: string[];
    dayOfMonth?: number;
    month?: number;
    firstDayOfWeek?: string;
    index?: string; // 'first' | 'second' | 'third' | 'fourth' | 'last'
  };
  range: {
    type: 'endDate' | 'noEnd' | 'numbered';
    startDate: string;
    endDate?: string;
    numberOfOccurrences?: number;
    recurrenceTimeZone?: string;
  };
}

/** Outlook event type classification */
export type OutlookEventType =
  | 'singleInstance'
  | 'seriesMaster'
  | 'occurrence'
  | 'exception';

/**
 * Enriched event ready for calendar-hub consumption.
 *
 * Produced by RecurrenceService.processEvent() — calendar-hub maps this
 * directly to CalendarEventEntity fields without needing to understand
 * the Outlook data model.
 */
export interface ProcessedOutlookEvent {
  externalId: string;
  eventType: OutlookEventType;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  subject: string;
  bodyPreview: string;
  location?: string;
  showAs?: string;
  changeKey?: string;
  seriesMasterId?: string;
  transactionId?: string;
  /** Only set on seriesMaster events */
  recurrenceRule?: RecurrenceRule;
  /** Only set on exception events (original time before user modified it) */
  originalStart?: { dateTime: string; timeZone: string };
}

/** Calculated expansion date range for recurring series */
export interface ExpansionWindow {
  startDate: Date;
  endDate: Date;
}

/** Options for expandRecurringSeries() */
export interface ExpandRecurringSeriesOptions {
  /** Currently stored occurrence external IDs (for stale detection) */
  existingExternalIds?: string[];
  /** Current expansion window end date (for window advancement) */
  existingWindowEnd?: Date;
}

/**
 * Full output of expanding a recurring series.
 *
 * Returned by RecurrenceService.expandRecurringSeries() — gives calendar-hub
 * everything it needs to persist the series master, upsert instances,
 * and clean up stale occurrences in a single orchestrated call.
 */
export interface RecurringEventExpansionResult {
  seriesMaster: ProcessedOutlookEvent;
  instances: ProcessedOutlookEvent[];
  expansionWindow: ExpansionWindow;
  /** External IDs that existed before but were not returned by this expansion */
  staleExternalIds: string[];
}
