import { Logger } from '@nestjs/common';
import { RecurrenceService } from './recurrence.service';
import { CalendarService } from './calendar.service';
import { Event } from '../../types';

const DAY_MS = 24 * 60 * 60 * 1000;

describe('RecurrenceService.sliceExpansionWindow', () => {
  const service = new RecurrenceService({} as CalendarService);

  it('covers a 5-year window in contiguous slices no longer than a year', () => {
    const startDate = new Date('2026-09-23T10:45:42.789Z');
    const endDate = new Date('2031-09-21T10:45:42.789Z');

    const slices = service.sliceExpansionWindow({ startDate, endDate });

    expect(slices.length).toBe(5);
    expect(slices[0].startDate).toEqual(startDate);
    expect(slices[slices.length - 1].endDate).toEqual(endDate);
    for (let i = 0; i < slices.length; i++) {
      const span = slices[i].endDate.getTime() - slices[i].startDate.getTime();
      expect(span).toBeLessThanOrEqual(RecurrenceService.MAX_EXPANSION_SLICE_DAYS * DAY_MS);
      if (i > 0) {
        expect(slices[i].startDate).toEqual(slices[i - 1].endDate);
      }
    }
  });

  it('returns a window shorter than the limit as a single slice', () => {
    const startDate = new Date('2026-01-06T00:00:00.000Z');
    const endDate = new Date('2026-09-23T10:45:42.789Z');

    expect(service.sliceExpansionWindow({ startDate, endDate })).toEqual([{ startDate, endDate }]);
  });

  it('returns no slices for an empty window', () => {
    const date = new Date('2026-09-23T00:00:00.000Z');

    expect(service.sliceExpansionWindow({ startDate: date, endDate: date })).toEqual([]);
  });
});

describe('RecurrenceService.expandRecurringSeries', () => {
  beforeEach(() => {
    jest.spyOn(Logger.prototype, 'log').mockImplementation(() => undefined);
  });

  afterEach(() => jest.restoreAllMocks());

  it('requests each slice separately and keeps one copy of an instance seen in two slices', async () => {
    const master: Event = {
      id: 'master',
      type: 'seriesMaster',
      recurrence: {
        pattern: { type: 'daily', interval: 1 },
        range: { type: 'noEnd', startDate: '2026-01-06' },
      },
    } as Event;
    const instance = (id: string): Event => ({ id, type: 'occurrence', seriesMasterId: 'master' }) as Event;

    const requestedRanges: Array<{ startDate: Date; endDate: Date }> = [];
    const calendarService = {
      getEventsBatch: jest.fn().mockResolvedValue([master]),
      getRecurringEventInstances: jest.fn(async function* (
        _seriesMasterId: string,
        _externalUserId: string,
        options: { startDate: Date; endDate: Date },
      ) {
        requestedRanges.push(options);
        // Every slice returns the shared boundary instance plus one of its own.
        yield [instance('boundary'), instance(`own-${requestedRanges.length}`)];
      }),
    } as unknown as CalendarService;

    const service = new RecurrenceService(calendarService);
    const result = await service.expandRecurringSeries('master', '1350');

    for (const range of requestedRanges) {
      const span = range.endDate.getTime() - range.startDate.getTime();
      expect(span).toBeLessThanOrEqual(RecurrenceService.MAX_EXPANSION_SLICE_DAYS * DAY_MS);
    }
    expect(requestedRanges.length).toBeGreaterThan(2);

    const ids = result.instances.map((i) => i.externalId);
    expect(ids.filter((id) => id === 'boundary')).toHaveLength(1);
    expect(new Set(ids).size).toBe(ids.length);
  });
});
