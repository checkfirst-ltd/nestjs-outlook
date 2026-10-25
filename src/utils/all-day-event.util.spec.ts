import {
  addCalendarDays,
  assertCalendarDate,
  buildAllDayDateTimeRange,
  calendarDateFromIso,
  normalizeAllDayEventTimes,
} from "./all-day-event.util";

describe("assertCalendarDate", () => {
  it("accepts a valid YYYY-MM-DD", () => {
    expect(() => {
      assertCalendarDate("2026-06-25");
    }).not.toThrow();
  });

  it("rejects invalid calendar dates", () => {
    expect(() => {
      assertCalendarDate("2026-02-30");
    }).toThrow(RangeError);
    expect(() => {
      assertCalendarDate("06-25-2026");
    }).toThrow(RangeError);
  });
});

describe("addCalendarDays", () => {
  it("rolls across month boundaries", () => {
    expect(addCalendarDays("2026-01-31", 1)).toBe("2026-02-01");
  });

  it("handles leap years", () => {
    expect(addCalendarDays("2024-02-28", 1)).toBe("2024-02-29");
    expect(addCalendarDays("2024-02-29", 1)).toBe("2024-03-01");
  });
});

describe("buildAllDayDateTimeRange", () => {
  it("builds a single-day range with exclusive end", () => {
    expect(
      buildAllDayDateTimeRange({
        startDate: "2026-06-25",
        timeZone: "Europe/London",
      })
    ).toEqual({
      start: { dateTime: "2026-06-25T00:00:00", timeZone: "Europe/London" },
      end: { dateTime: "2026-06-26T00:00:00", timeZone: "Europe/London" },
    });
  });

  it("builds a multi-day inclusive range", () => {
    expect(
      buildAllDayDateTimeRange({
        startDate: "2026-06-25",
        endDateInclusive: "2026-06-27",
        timeZone: "UTC",
      })
    ).toEqual({
      start: { dateTime: "2026-06-25T00:00:00", timeZone: "UTC" },
      end: { dateTime: "2026-06-28T00:00:00", timeZone: "UTC" },
    });
  });

  it("accepts ISO strings for start and end", () => {
    expect(
      buildAllDayDateTimeRange({
        startDate: "2026-06-25T00:00:00",
        endDateInclusive: "2026-06-27",
        timeZone: "UTC",
      })
    ).toEqual({
      start: { dateTime: "2026-06-25T00:00:00", timeZone: "UTC" },
      end: { dateTime: "2026-06-28T00:00:00", timeZone: "UTC" },
    });
  });

  it("rejects end before start", () => {
    expect(() =>
      buildAllDayDateTimeRange({
        startDate: "2026-06-27",
        endDateInclusive: "2026-06-25",
        timeZone: "UTC",
      })
    ).toThrow(RangeError);
  });
});

describe("normalizeAllDayEventTimes", () => {
  it("strips Z suffix and non-midnight times for Graph all-day create", () => {
    const normalized = normalizeAllDayEventTimes({
      isAllDay: true,
      subject: "Free block",
      start: {
        dateTime: "2026-10-21T14:30:00.000Z",
        timeZone: "UTC",
      },
      end: {
        dateTime: "2026-10-22T00:00:00.000Z",
        timeZone: "UTC",
      },
    });

    expect(normalized.start).toEqual({
      dateTime: "2026-10-21T00:00:00",
      timeZone: "UTC",
    });
    expect(normalized.end).toEqual({
      dateTime: "2026-10-22T00:00:00",
      timeZone: "UTC",
    });
  });
});

describe("calendarDateFromIso", () => {
  it("accepts YYYY-MM-DD", () => {
    expect(calendarDateFromIso("2026-06-25")).toBe("2026-06-25");
  });

  it("strips the date from an ISO datetime", () => {
    expect(calendarDateFromIso("2026-06-25T14:30:00Z", "start")).toBe(
      "2026-06-25"
    );
  });
});
