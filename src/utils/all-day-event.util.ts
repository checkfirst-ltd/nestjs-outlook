import { DateTimeTimeZone, Event } from "../types";

const CALENDAR_DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

/**
 * Extracts a YYYY-MM-DD calendar date from a date-only or ISO 8601 string.
 */
export function calendarDateFromIso(iso: string, fieldName = "date"): string {
  const trimmed = iso.trim();
  const datePart =
    trimmed.length === 10 && CALENDAR_DATE_RE.test(trimmed)
      ? trimmed
      : trimmed.match(/^(\d{4}-\d{2}-\d{2})/)?.[1];

  if (!datePart) {
    throw new RangeError(
      `${fieldName} must be YYYY-MM-DD or an ISO 8601 string with a calendar date`
    );
  }

  assertCalendarDate(datePart, fieldName);
  return datePart;
}

/**
 * Validates a calendar date string (YYYY-MM-DD) and throws if invalid.
 */
export function assertCalendarDate(date: string, fieldName = "date"): void {
  if (!CALENDAR_DATE_RE.test(date)) {
    throw new RangeError(`${fieldName} must be YYYY-MM-DD, got ${date}`);
  }

  const [y, m, d] = date.split("-").map(Number);
  const parsed = new Date(Date.UTC(y, m - 1, d));
  if (
    parsed.getUTCFullYear() !== y ||
    parsed.getUTCMonth() !== m - 1 ||
    parsed.getUTCDate() !== d
  ) {
    throw new RangeError(`${fieldName} is not a valid calendar date: ${date}`);
  }
}

/**
 * Adds calendar days to a YYYY-MM-DD date (UTC calendar arithmetic).
 */
export function addCalendarDays(date: string, days: number): string {
  assertCalendarDate(date);
  const [y, m, d] = date.split("-").map(Number);
  const next = new Date(Date.UTC(y, m - 1, d + days));
  const yy = next.getUTCFullYear();
  const mm = String(next.getUTCMonth() + 1).padStart(2, "0");
  const dd = String(next.getUTCDate()).padStart(2, "0");
  return `${yy}-${mm}-${dd}`;
}

export interface AllDayDateTimeRangeInput {
  /** First included day — YYYY-MM-DD or ISO 8601 (date portion is used). */
  startDate: string;
  /** Last included day — YYYY-MM-DD or ISO; defaults to `startDate`. */
  endDateInclusive?: string;
  /** IANA time zone for midnight boundaries on start/end (e.g. `Europe/London`). */
  timeZone: string;
}

/**
 * Builds Microsoft Graph `start` / `end` for an all-day event.
 * Graph uses an exclusive end at midnight on the day after the last included day.
 * Set `isAllDay` and `showAs` on the event body yourself.
 *
 * @see https://learn.microsoft.com/en-us/graph/api/resources/datetimetimezone
 */
export function buildAllDayDateTimeRange(
  input: AllDayDateTimeRangeInput
): { start: DateTimeTimeZone; end: DateTimeTimeZone } {
  const startDate = calendarDateFromIso(input.startDate, "startDate");
  const endInclusive = input.endDateInclusive
    ? calendarDateFromIso(input.endDateInclusive, "endDateInclusive")
    : startDate;

  if (endInclusive < startDate) {
    throw new RangeError(
      `endDateInclusive (${endInclusive}) must not be before startDate (${startDate})`
    );
  }

  const endExclusive = addCalendarDays(endInclusive, 1);
  return {
    start: {
      dateTime: formatAllDayMidnightDateTime(startDate),
      timeZone: input.timeZone,
    },
    end: {
      dateTime: formatAllDayMidnightDateTime(endExclusive),
      timeZone: input.timeZone,
    },
  };
}

/** Graph all-day `dateTime` must be `YYYY-MM-DDT00:00:00` without a `Z` offset suffix. */
export function formatAllDayMidnightDateTime(calendarDate: string): string {
  assertCalendarDate(calendarDate);
  return `${calendarDate}T00:00:00`;
}

function isMidnightDateTimeValue(dateTime: string): boolean {
  return /T00:00:00(\.\d+)?(Z)?$/i.test(dateTime.trim());
}

/**
 * Rebuilds `start` / `end` for an all-day Graph event (midnight, same zone, exclusive end).
 * Use when callers pass ISO strings with `Z`, offsets, or non-midnight clock times.
 */
export function normalizeAllDayEventTimes(
  event: Partial<Event>
): Partial<Event> {
  if (!event.isAllDay) {
    return event;
  }

  const timeZone =
    event.start?.timeZone ?? event.end?.timeZone ?? "UTC";

  if (!event.start?.dateTime) {
    return event;
  }

  const startDate = calendarDateFromIso(event.start.dateTime, "start");

  let endDateInclusive = startDate;
  if (event.end?.dateTime) {
    const endRaw = event.end.dateTime;
    const endDate = calendarDateFromIso(endRaw, "end");
    if (isMidnightDateTimeValue(endRaw) && endDate > startDate) {
      endDateInclusive = addCalendarDays(endDate, -1);
    } else {
      endDateInclusive = endDate;
    }
  }

  const range = buildAllDayDateTimeRange({
    startDate,
    endDateInclusive,
    timeZone,
  });

  return {
    ...event,
    start: range.start,
    end: range.end,
  };
}
