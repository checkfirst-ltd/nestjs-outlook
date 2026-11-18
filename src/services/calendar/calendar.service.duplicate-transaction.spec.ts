import { Logger } from '@nestjs/common';
import { Client } from '@microsoft/microsoft-graph-client';
import axios from 'axios';
import { CalendarService } from './calendar.service';
import { DuplicateTransactionIdError } from '../../errors/duplicate-transaction-id.error';
import { FakeGraphMailbox } from '../../mock/fake-graph-mailbox';
import type { Event } from '../../types';

jest.mock('@microsoft/microsoft-graph-client', () => ({ Client: { init: jest.fn() } }));
jest.mock('axios');

/**
 * Outlook showed auditors two copies of one appointment. These tests drive the real
 * CalendarService (its retries, rate limiting and error handling) against a simulated mailbox
 * that answers the way Graph answered in production:
 *
 *   201                          the event is created
 *   502 after writing the event  the answer is lost, the event is not
 *   400 ErrorDuplicateTransactionId with @event.existingEventId
 *                                a create reused a transactionId; Graph names the holder
 *
 * Before the fix the 400 was retried seven times and then failed, so the event that already
 * existed was never linked and later surfaced as a duplicate.
 */
describe('CalendarService: creating an event whose transactionId is already used', () => {
  const user = '1708';
  const calendarId = 'calendar-1';
  const transactionId = 'calendar-hub-with-app-id:app-1-scheduleai-activity-109134';

  let mailbox: FakeGraphMailbox;
  let service: CalendarService;

  const payload = (overrides: Partial<Event> = {}): Partial<Event> => ({
    subject: 'Audit 50051214 (Stage 2)',
    start: { dateTime: '2027-01-07T01:00:00', timeZone: 'UTC' },
    end: { dateTime: '2027-01-07T09:00:00', timeZone: 'UTC' },
    showAs: 'busy',
    transactionId,
    ...overrides,
  });

  /** Run a call to completion, letting retry backoffs elapse in fake time. */
  async function settle<T>(call: Promise<T>): Promise<T> {
    let outcome: { ok: true; value: T } | { ok: false; error: unknown } | undefined;
    void call.then(
      (value) => { outcome = { ok: true, value }; },
      (error: unknown) => { outcome = { ok: false, error }; },
    );
    for (let i = 0; !outcome && i < 600; i++) {
      await jest.advanceTimersByTimeAsync(1_000);
    }
    if (!outcome) throw new Error('call did not settle within 10 minutes of fake time');
    if (!outcome.ok) throw outcome.error;
    return outcome.value;
  }

  beforeEach(() => {
    jest.useFakeTimers();
    jest.spyOn(Logger.prototype, 'log').mockImplementation(() => undefined);
    jest.spyOn(Logger.prototype, 'warn').mockImplementation(() => undefined);
    jest.spyOn(Logger.prototype, 'error').mockImplementation(() => undefined);
    jest.spyOn(Logger.prototype, 'debug').mockImplementation(() => undefined);

    mailbox = new FakeGraphMailbox();
    (Client.init as jest.Mock).mockImplementation(() => mailbox.sdkClient());
    (axios.post as jest.Mock).mockImplementation((url: string, body: never) => mailbox.axiosPost(url, body));

    const rateLimiter = {
      acquirePermit: jest.fn().mockResolvedValue(undefined),
      releasePermit: jest.fn(),
      recordSuccess: jest.fn().mockResolvedValue(undefined),
      record503Failure: jest.fn().mockResolvedValue(undefined),
      handleRateLimitResponse: jest.fn().mockResolvedValue(undefined),
    };

    service = new CalendarService(
      { getUserAccessToken: jest.fn().mockResolvedValue('token') } as never, // microsoftAuthService
      {} as never, // eventEmitter
      {} as never, // microsoftConfig
      {} as never, // deltaLinkRepository
      { findOne: jest.fn().mockResolvedValue(null) } as never, // microsoftUserRepository → delegated /me
      {} as never, // deltaSyncService
      {} as never, // userIdConverter
      {} as never, // subscriptionService
      rateLimiter as never,
      null,
    );
  });

  afterEach(() => {
    jest.useRealTimers();
    jest.restoreAllMocks();
  });

  describe('createEvent', () => {
    it('201: creates the event once', async () => {
      const { event, adoptedExisting } = await settle(service.createEvent(payload(), user, calendarId));

      expect(adoptedExisting).toBeUndefined();
      expect(event.transactionId).toBe(transactionId);
      expect(mailbox.eventsWithTransactionId(transactionId)).toHaveLength(1);
      expect(mailbox.count('POST')).toBe(1);
    });

    it('400 duplicate: adopts the event Graph names instead of retrying, and leaves one copy', async () => {
      // An earlier attempt wrote the event but its id was never stored.
      const orphan = mailbox.seed({
        subject: 'Audit 50051214 (old title)',
        start: { dateTime: '2027-01-06T01:00:00', timeZone: 'UTC' },
        end: { dateTime: '2027-01-06T09:00:00', timeZone: 'UTC' },
        transactionId,
      });

      const { event, adoptedExisting } = await settle(service.createEvent(payload(), user, calendarId));

      expect(adoptedExisting).toBe(true);
      expect(event.id).toBe(orphan.id);
      expect(event.transactionId).toBe(transactionId);
      // The adopted event now says what the appointment says.
      expect(event.subject).toBe('Audit 50051214 (Stage 2)');
      expect(event.start?.dateTime).toBe('2027-01-07T01:00:00');
      // One POST: the duplicate is not retried. One PATCH: the adoption.
      expect(mailbox.count('POST')).toBe(1);
      expect(mailbox.count('PATCH', `/me/events/${orphan.id}`)).toBe(1);
      expect(mailbox.eventsWithTransactionId(transactionId)).toHaveLength(1);
    });

    it('502 after the write: the retry meets the duplicate and adopts the event the first try made', async () => {
      mailbox.loseAnswersOfNextCreates(1);

      const { event, adoptedExisting } = await settle(service.createEvent(payload(), user, calendarId));

      expect(adoptedExisting).toBe(true);
      expect(mailbox.count('POST')).toBe(2);
      expect(mailbox.eventsWithTransactionId(transactionId)).toEqual([
        expect.objectContaining({ id: event.id }),
      ]);
    });

    it('400 duplicate whose holder is gone: fails at once with DuplicateTransactionIdError', async () => {
      const orphan = mailbox.seed({
        subject: 'gone',
        start: { dateTime: '2027-01-07T01:00:00', timeZone: 'UTC' },
        end: { dateTime: '2027-01-07T09:00:00', timeZone: 'UTC' },
        transactionId,
      });
      mailbox.events.delete(orphan.id);

      const error = await settle(service.createEvent(payload(), user, calendarId)).catch((e: unknown) => e);

      expect(error).toBeInstanceOf(DuplicateTransactionIdError);
      expect(error).toMatchObject({ transactionId, existingEventId: orphan.id });
      expect(mailbox.count('POST')).toBe(1);
      expect(mailbox.events.size).toBe(0);
    });

    it('400 duplicate that names no holder: fails at once with DuplicateTransactionIdError', async () => {
      mailbox.seed({
        subject: 'orphan',
        start: { dateTime: '2027-01-07T01:00:00', timeZone: 'UTC' },
        end: { dateTime: '2027-01-07T09:00:00', timeZone: 'UTC' },
        transactionId,
      });
      mailbox.omitExistingEventIdOnDuplicates();

      const error = await settle(service.createEvent(payload(), user, calendarId)).catch((e: unknown) => e);

      expect(error).toBeInstanceOf(DuplicateTransactionIdError);
      expect(error).toMatchObject({ existingEventId: null });
      expect(mailbox.count('POST')).toBe(1);
      expect(mailbox.count('PATCH')).toBe(0);
    });
  });

  describe('createBatchEvents', () => {
    const txn = (n: number) => `calendar-hub-with-app-id:app-1-scheduleai-activity-${n}`;

    it('adopts the item whose transactionId is taken, and maps every answer to its own request', async () => {
      mailbox.answerBatchesInReverse = true;
      const orphan = mailbox.seed({
        subject: 'orphan of 2',
        start: { dateTime: '2027-01-05T01:00:00', timeZone: 'UTC' },
        end: { dateTime: '2027-01-05T09:00:00', timeZone: 'UTC' },
        transactionId: txn(2),
      });
      const events = [1, 2, 3].map((n) => payload({ subject: `Appointment ${n}`, transactionId: txn(n) }));

      const results = await settle(service.createBatchEvents(events, user, calendarId));

      expect(results.map((r) => r.index)).toEqual([0, 1, 2]);
      expect(results.every((r) => r.success)).toBe(true);
      // Each result carries the event of its own request, though Graph answered last-first.
      expect(results.map((r) => r.event?.transactionId)).toEqual([txn(1), txn(2), txn(3)]);
      expect(results[1]).toMatchObject({ adoptedExisting: true, event: { id: orphan.id, subject: 'Appointment 2' } });
      for (const n of [1, 2, 3]) {
        expect(mailbox.eventsWithTransactionId(txn(n))).toHaveLength(1);
      }
    });
  });

  describe('deleteBatchEvents', () => {
    it('maps every answer to its own event when Graph answers out of order', async () => {
      mailbox.answerBatchesInReverse = true;
      const kept = mailbox.seed({
        subject: 'denied',
        start: { dateTime: '2027-01-05T01:00:00', timeZone: 'UTC' },
        end: { dateTime: '2027-01-05T09:00:00', timeZone: 'UTC' },
      });
      const removed = mailbox.seed({
        subject: 'deleted',
        start: { dateTime: '2027-01-06T01:00:00', timeZone: 'UTC' },
        end: { dateTime: '2027-01-06T09:00:00', timeZone: 'UTC' },
      });
      mailbox.forceAnswer('DELETE', kept.id, {
        status: 403,
        body: { error: { code: 'ErrorAccessDenied', message: 'Access is denied.' } },
      });

      const results = await settle(service.deleteBatchEvents([kept.id, removed.id], user, calendarId));

      expect(results).toEqual(
        expect.arrayContaining([
          expect.objectContaining({ id: kept.id, success: false }),
          { id: removed.id, success: true },
        ]),
      );
      expect(mailbox.events.has(kept.id)).toBe(true);
      expect(mailbox.events.has(removed.id)).toBe(false);
    });
  });

  describe('findEventsByTransactionId', () => {
    it('finds the event by transactionId across calendar-view pages', async () => {
      for (let h = 0; h < 150; h++) {
        const at = new Date(Date.UTC(2027, 0, 7, 0, 0) + h * 60_000).toISOString().slice(0, 19);
        mailbox.seed({ subject: `busy ${h}`, start: { dateTime: at, timeZone: 'UTC' }, end: { dateTime: at, timeZone: 'UTC' } });
      }
      const target = mailbox.seed({
        subject: 'the appointment',
        start: { dateTime: '2027-01-07T23:00:00', timeZone: 'UTC' },
        end: { dateTime: '2027-01-07T23:30:00', timeZone: 'UTC' },
        transactionId,
      });

      const found = await settle(
        service.findEventsByTransactionId(user, transactionId, {
          start: '2027-01-07T01:00:00',
          end: '2027-01-07T09:00:00',
        }),
      );

      expect(found.map((e) => e.id)).toEqual([target.id]);
      expect(mailbox.count('GET', 'calendarView')).toBe(2);
    });

    it('returns nothing when no event in the window has the transactionId', async () => {
      mailbox.seed({
        subject: 'outside the window',
        start: { dateTime: '2027-03-01T01:00:00', timeZone: 'UTC' },
        end: { dateTime: '2027-03-01T09:00:00', timeZone: 'UTC' },
        transactionId,
      });

      const found = await settle(
        service.findEventsByTransactionId(user, transactionId, {
          start: '2027-01-07T01:00:00',
          end: '2027-01-07T09:00:00',
        }),
      );

      expect(found).toEqual([]);
    });
  });
});
