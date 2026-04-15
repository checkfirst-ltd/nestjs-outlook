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
   * Full orchestration: fetch the series master and expand its instances.
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

    // 2. Calculate expansion windows (past + future halves)
    const expansionWindows = this.calculateExpansionWindow(
      seriesMaster.recurrenceRule,
    );

    // 3. Fetch and process all instances within each window
    const instances: ProcessedOutlookEvent[] = [];

    for (const window of expansionWindows) {
      this.logger.log(
        `[expandRecurringSeries] Window: ${window.startDate.toISOString()} → ${window.endDate.toISOString()}`,
      );

      for await (const batch of this.calendarService.getRecurringEventInstances(
        seriesMasterId,
        externalUserId,
        {
          startDate: window.startDate,
          endDate: window.endDate,
          batchSize: 100,
        },
      )) {
        for (const event of batch) {
          instances.push(this.processEvent(event));
        }
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
      expansionWindow: expansionWindows,
      staleExternalIds,
    };
  }

  /**
   * Calculate the date ranges for expanding recurring event instances.
   *
   * Graph's calendarView caps a single expansion request at a 5-year span,
   * so we split the work into two halves — past (now-5y → now) and future
   * (now → now+5y) — each clamped by the series' own range when known.
   *
   * A 2-day safety margin is subtracted from the far edges to stay inside
   * Outlook's hard 5-year limit.
   */
  calculateExpansionWindow(
    recurrenceRule?: RecurrenceRule,
  ): ExpansionWindow[] {
    const now = new Date();

    const fiveYearsAgo = new Date(now);
    fiveYearsAgo.setFullYear(fiveYearsAgo.getFullYear() - 5);
    fiveYearsAgo.setDate(fiveYearsAgo.getDate() + 2);

    const fiveYearsAhead = new Date(now);
    fiveYearsAhead.setFullYear(fiveYearsAhead.getFullYear() + 5);
    fiveYearsAhead.setDate(fiveYearsAhead.getDate() - 2);

    const seriesStart = recurrenceRule?.range.startDate
      ? new Date(recurrenceRule.range.startDate)
      : undefined;
    const seriesEnd =
      recurrenceRule?.range.type === 'endDate' && recurrenceRule.range.endDate
        ? new Date(recurrenceRule.range.endDate)
        : undefined;

    const windows: ExpansionWindow[] = [];

    // Past half: max(now-5y, seriesStart) → now
    const pastStart =
      seriesStart && seriesStart > fiveYearsAgo ? seriesStart : fiveYearsAgo;
    if (pastStart < now) {
      windows.push({ startDate: pastStart, endDate: new Date(now) });
    }

    // Future half: now → min(now+5y, seriesEnd)
    const futureEnd =
      seriesEnd && seriesEnd < fiveYearsAhead ? seriesEnd : fiveYearsAhead;
    const futureStart =
      seriesStart && seriesStart > now ? seriesStart : new Date(now);
    if (futureEnd > futureStart) {
      windows.push({ startDate: futureStart, endDate: futureEnd });
    }

    return windows;
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
