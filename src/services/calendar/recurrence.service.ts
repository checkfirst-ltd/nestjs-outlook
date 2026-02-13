import { Injectable, Logger } from '@nestjs/common';
import { PatternedRecurrence } from '@microsoft/microsoft-graph-types';
import { Event } from '../../types';
import { CalendarService } from './calendar.service';
import {
  RecurrenceRule,
  OutlookEventType,
  ProcessedOutlookEvent,
  ExpansionWindow,
  ExpandRecurringSeriesOptions,
  RecurringEventExpansionResult,
} from '../../interfaces/recurrence.interfaces';

@Injectable()
export class RecurrenceService {
  private readonly logger = new Logger(RecurrenceService.name);

  constructor(private readonly calendarService: CalendarService) {}

  /**
   * Map a raw Microsoft Graph Event to an enriched ProcessedOutlookEvent.
   *
   * Extracts recurrence metadata that the raw Event carries but that
   * calendar-hub previously ignored: eventType, recurrenceRule, originalStart.
   */
  processEvent(event: Event): ProcessedOutlookEvent {
    const eventType: OutlookEventType = event.type ?? 'singleInstance';

    const processed: ProcessedOutlookEvent = {
      externalId: event.id ?? '',
      eventType,
      start: {
        dateTime: event.start?.dateTime ?? '',
        timeZone: event.start?.timeZone ?? '',
      },
      end: {
        dateTime: event.end?.dateTime ?? '',
        timeZone: event.end?.timeZone ?? '',
      },
      subject: event.subject ?? '',
      bodyPreview: event.bodyPreview ?? '',
      location: event.location?.displayName ?? undefined,
      showAs: event.showAs ?? undefined,
      changeKey: event.changeKey ?? undefined,
      seriesMasterId: event.seriesMasterId ?? undefined,
      transactionId: event.transactionId ?? undefined,
    };

    // Attach recurrence rule only for series masters
    if (eventType === 'seriesMaster' && event.recurrence) {
      processed.recurrenceRule = this.extractRecurrenceRule(event.recurrence);
    }

    // Attach original start only for exceptions
    if (eventType === 'exception' && event.originalStart) {
      processed.originalStart = {
        dateTime: event.originalStart,
        timeZone:
          event.originalStartTimeZone ?? event.start?.timeZone ?? '',
      };
    }

    return processed;
  }

  /**
   * Full orchestration: fetch the series master, expand its instances,
   * and detect stale occurrences — all in one call.
   *
   * Returns everything calendar-hub needs to persist the series:
   * - seriesMaster (with recurrenceRule for metadata storage)
   * - instances (enriched occurrences/exceptions)
   * - expansionWindow (for tracking how far ahead we've expanded)
   * - staleExternalIds (occurrences to soft-delete)
   */
  async expandRecurringSeries(
    seriesMasterId: string,
    externalUserId: string,
    options?: ExpandRecurringSeriesOptions,
  ): Promise<RecurringEventExpansionResult> {
    this.logger.log(
      `[expandRecurringSeries] Expanding series ${seriesMasterId} for user ${externalUserId}`,
    );

    // 1. Fetch the series master event
    const masterEvents = await this.calendarService.getEventsBatch(
      [seriesMasterId],
      externalUserId,
    );

    if (masterEvents.length === 0) {
      throw new Error(
        `Series master ${seriesMasterId} not found in Outlook for user ${externalUserId}`,
      );
    }

    const seriesMaster = this.processEvent(masterEvents[0]);

    // 2. Calculate expansion window
    const expansionWindow = this.calculateExpansionWindow(
      seriesMaster.recurrenceRule,
      options?.existingWindowEnd,
    );

    this.logger.log(
      `[expandRecurringSeries] Window: ${expansionWindow.startDate.toISOString()} → ${expansionWindow.endDate.toISOString()}`,
    );

    // 3. Fetch and process all instances within the window
    const instances: ProcessedOutlookEvent[] = [];

    for await (const batch of this.calendarService.getRecurringEventInstances(
      seriesMasterId,
      externalUserId,
      {
        startDate: expansionWindow.startDate,
        endDate: expansionWindow.endDate,
        batchSize: 100,
      },
    )) {
      for (const event of batch) {
        instances.push(this.processEvent(event));
      }
    }

    this.logger.log(
      `[expandRecurringSeries] Fetched ${instances.length} instances for series ${seriesMasterId}`,
    );

    // 4. Detect stale occurrences
    const staleExternalIds = options?.existingExternalIds
      ? this.detectStaleOccurrences(
          instances.map((i) => i.externalId),
          options.existingExternalIds,
        )
      : [];

    if (staleExternalIds.length > 0) {
      this.logger.log(
        `[expandRecurringSeries] Detected ${staleExternalIds.length} stale occurrences`,
      );
    }

    return {
      seriesMaster,
      instances,
      expansionWindow,
      staleExternalIds,
    };
  }

  /**
   * Calculate the date range for expanding recurring event instances.
   *
   * Strategy:
   * - Start date: always 1 month before now
   * - End date depends on series type:
   *   - Finite series (endDate range): use the series end date, capped at 5 years
   *   - Advancing existing window: extend 6 months from current window end
   *   - New infinite series: 6 months from now
   */
  calculateExpansionWindow(
    recurrenceRule?: RecurrenceRule,
    existingWindowEnd?: Date,
  ): ExpansionWindow {
    const now = new Date();

    const startDate = new Date(now);
    startDate.setMonth(startDate.getMonth() - 1);

    let endDate: Date;

    if (
      recurrenceRule?.range.type === 'endDate' &&
      recurrenceRule.range.endDate
    ) {
      // Series has a defined end — use it, capped at 5 years
      const seriesEnd = new Date(recurrenceRule.range.endDate);
      const maxEnd = new Date(now);
      maxEnd.setFullYear(maxEnd.getFullYear() + 5);
      endDate = seriesEnd < maxEnd ? seriesEnd : maxEnd;
    } else if (existingWindowEnd) {
      // Advancing an existing window — extend 6 months from current end
      endDate = new Date(existingWindowEnd);
      endDate.setMonth(endDate.getMonth() + 6);
    } else {
      // New infinite series — start with 6 months ahead
      endDate = new Date(now);
      endDate.setMonth(endDate.getMonth() + 6);
    }

    return { startDate, endDate };
  }

  /**
   * Return existing external IDs that were NOT returned by the latest expansion.
   * These represent deleted or removed occurrences that should be soft-deleted.
   */
  detectStaleOccurrences(
    fetchedExternalIds: string[],
    existingExternalIds: string[],
  ): string[] {
    const fetchedSet = new Set(fetchedExternalIds);
    return existingExternalIds.filter((id) => !fetchedSet.has(id));
  }

  /**
   * Check if an event is a recurring series master.
   */
  isSeriesMaster(event: Event): boolean {
    return event.type === 'seriesMaster' || event.recurrence != null;
  }

  /**
   * Extract a clean RecurrenceRule from Microsoft Graph's PatternedRecurrence.
   */
  private extractRecurrenceRule(
    recurrence: PatternedRecurrence,
  ): RecurrenceRule {
    const pattern = recurrence.pattern;
    const range = recurrence.range;

    return {
      pattern: {
        type: pattern?.type ?? 'daily',
        interval: pattern?.interval ?? 1,
        daysOfWeek: pattern?.daysOfWeek ?? undefined,
        dayOfMonth: pattern?.dayOfMonth ?? undefined,
        month: pattern?.month ?? undefined,
        firstDayOfWeek: pattern?.firstDayOfWeek ?? undefined,
        index: pattern?.index ?? undefined,
      },
      range: {
        type: range?.type ?? 'noEnd',
        startDate: range?.startDate ?? '',
        endDate: range?.endDate ?? undefined,
        numberOfOccurrences: range?.numberOfOccurrences ?? undefined,
        recurrenceTimeZone: range?.recurrenceTimeZone ?? undefined,
      },
    };
  }
}
