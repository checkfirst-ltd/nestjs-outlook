import type { Event } from '../types';

/**
 * A single Outlook mailbox behind a fake Microsoft Graph, for tests that drive the real
 * CalendarService (its retry, rate-limit and error paths) without a network.
 *
 * It keeps the Graph behaviour the duplicate-event bugs depend on:
 * - a create whose transactionId is already used answers 400 `ErrorDuplicateTransactionId`
 *   and names the holder in `@event.existingEventId`, exactly as production logs show
 * - a create can be told to write the event and then lose the answer as a 502
 * - `$batch` can answer its items in reverse order, which Graph is allowed to do
 * - the calendar view pages with `$top`/`$skip` and an `@odata.nextLink`
 *
 * The SDK side is reached through {@link sdkClient} (stand-in for `Client.init`) and the
 * `$batch` side through {@link axiosPost} (stand-in for `axios.post`).
 */

type GraphErrorBody = Record<string, unknown> & { code: string; message: string };

export interface RecordedRequest {
  method: string;
  path: string;
  body?: unknown;
  viaBatch: boolean;
}

interface Answer {
  status: number;
  body?: unknown;
}

interface BatchRequestItem {
  id: string;
  method: string;
  url: string;
  body?: unknown;
}

const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0';

const DUPLICATE_MESSAGE =
  "Your request can't be completed. The TransactionId specified in the request has already been used to create a different event.";

/** The answer Graph gives when a create reuses a transactionId, copied from production. */
export function duplicateTransactionIdBody(existingEventId: string | null): GraphErrorBody {
  return {
    code: 'ErrorDuplicateTransactionId',
    message: DUPLICATE_MESSAGE,
    innererror: {
      message: DUPLICATE_MESSAGE,
      type: 'Microsoft.Exchange.Services.OData.ODataResponseException',
      internalexception: {
        message: DUPLICATE_MESSAGE,
        type: 'Microsoft.Exchange.Entities.Calendaring.ClientIdAlreadyInUseException',
      },
    },
    ...(existingEventId ? { '@event.existingEventId': existingEventId } : {}),
  };
}

/** What the Graph SDK throws: status, code and the JSON of Graph's error object in `body`. */
function sdkError(status: number, error: GraphErrorBody): Error {
  return Object.assign(new Error(error.message), {
    statusCode: status,
    code: error.code,
    body: JSON.stringify(error),
  });
}

function toUtcMs(dateTime: string | null | undefined): number {
  if (!dateTime) return Number.NaN;
  return new Date(/(?:Z|[+-]\d{2}:?\d{2})$/i.test(dateTime) ? dateTime : `${dateTime}Z`).getTime();
}

export class FakeGraphMailbox {
  /** Graph event id → event: what the auditor sees in Outlook. */
  readonly events = new Map<string, Event>();
  /** Every request that reached the mailbox, in order. */
  readonly requests: RecordedRequest[] = [];
  /**
   * transactionId → id of the event that claimed it. A claim outlives its event, so a test can
   * have Graph name a holder that is already gone.
   */
  private readonly transactionIds = new Map<string, string>();
  private seq = 0;
  private lostCreateAnswers = 0;
  private duplicatesWithoutHolder = false;
  private readonly forcedAnswers: Array<{ method: string; pathIncludes: string; answer: Answer }> = [];
  /** Answer `$batch` items last-first. Graph does not promise to keep the request order. */
  answerBatchesInReverse = false;

  // ── Test controls ──────────────────────────────────────────────────────────

  /** Put an event in the mailbox, e.g. an orphan left by an earlier attempt. */
  seed(
    event: Partial<Event> & { subject: string; start: Event['start']; end: Event['end'] },
  ): Event & { id: string } {
    const id = event.id ?? this.nextId();
    const stored = { ...event, id, changeKey: `ck-${id}-0` } as Event & { id: string };
    this.events.set(id, stored);
    if (stored.transactionId) {
      this.transactionIds.set(stored.transactionId, id);
    }
    return stored;
  }

  /** The next `count` creates are written, then answered with a 502 as if the answer was lost. */
  loseAnswersOfNextCreates(count: number): void {
    this.lostCreateAnswers += count;
  }

  /** Answer duplicate transactionIds without naming the event that holds them. */
  omitExistingEventIdOnDuplicates(): void {
    this.duplicatesWithoutHolder = true;
  }

  /** Answer every request matching `method` and `pathIncludes` with `answer`. */
  forceAnswer(method: string, pathIncludes: string, answer: Answer): void {
    this.forcedAnswers.push({ method: method.toUpperCase(), pathIncludes, answer });
  }

  /** The events that carry a transactionId: copies of one appointment show up here. */
  eventsWithTransactionId(transactionId: string): Event[] {
    return [...this.events.values()].filter((e) => e.transactionId === transactionId);
  }

  count(method: string, pathIncludes = ''): number {
    return this.requests.filter(
      (r) => r.method === method.toUpperCase() && r.path.includes(pathIncludes),
    ).length;
  }

  // ── Stand-ins for the transports CalendarService uses ──────────────────────

  /** Stand-in for the object `Client.init` returns. */
  sdkClient(): { api: (path: string) => unknown } {
    return {
      api: (path: string) => {
        const request = {
          header: () => request,
          headers: () => request,
          select: () => request,
          top: () => request,
          query: () => request,
          get: () => this.sdkCall('GET', path),
          post: (body: unknown) => this.sdkCall('POST', path, body),
          patch: (body: unknown) => this.sdkCall('PATCH', path, body),
          delete: () => this.sdkCall('DELETE', path),
        };
        return request;
      },
    };
  }

  /** Stand-in for `axios.post`, which CalendarService uses only for `$batch`. */
  axiosPost(url: string, payload: { requests: BatchRequestItem[] }): Promise<{ status: number; data: unknown }> {
    if (!url.endsWith('/$batch')) {
      return Promise.reject(new Error(`FakeGraphMailbox: unexpected axios.post to ${url}`));
    }
    const responses = payload.requests.map((item) => {
      const answer = this.handle(item.method.toUpperCase(), item.url, item.body, true);
      return { id: item.id, status: answer.status, body: answer.body };
    });
    if (this.answerBatchesInReverse) {
      responses.reverse();
    }
    return Promise.resolve({ status: 200, data: { responses } });
  }

  // ── Graph behaviour ────────────────────────────────────────────────────────

  private sdkCall(method: string, path: string, body?: unknown): Promise<unknown> {
    const answer = this.handle(method, path, body, false);
    if (answer.status >= 400) {
      const error = (answer.body as { error: GraphErrorBody }).error;
      return Promise.reject(sdkError(answer.status, error));
    }
    return Promise.resolve(answer.body);
  }

  private handle(method: string, rawPath: string, body: unknown, viaBatch: boolean): Answer {
    const path = rawPath.startsWith(GRAPH_ROOT) ? rawPath.slice(GRAPH_ROOT.length) : rawPath;
    this.requests.push({ method, path, body, viaBatch });

    const forced = this.forcedAnswers.find(
      (f) => f.method === method && path.includes(f.pathIncludes),
    );
    if (forced) {
      return forced.answer;
    }

    const [route, query = ''] = path.split('?');
    const segments = route.split('/').filter(Boolean);

    if (segments.includes('calendarView') && method === 'GET') {
      return this.calendarView(new URLSearchParams(query));
    }
    if (method === 'POST' && segments[segments.length - 1] === 'events') {
      return this.create(body as Partial<Event>);
    }

    const eventsAt = segments.lastIndexOf('events');
    const eventId = eventsAt >= 0 ? segments[eventsAt + 1] : undefined;
    if (!eventId) {
      return this.error(400, 'BadRequest', `FakeGraphMailbox does not serve ${method} ${path}`);
    }
    const existing = this.events.get(eventId);
    if (!existing) {
      return this.error(404, 'ErrorItemNotFound', 'The specified object was not found in the store.');
    }

    switch (method) {
      case 'GET':
        return { status: 200, body: existing };
      case 'PATCH': {
        // transactionId cannot change once an event has one.
        const updates = { ...(body as Partial<Event>) };
        delete updates.transactionId;
        const updated = {
          ...existing,
          ...updates,
          changeKey: `ck-${eventId}-${++this.seq}`,
        } as Event;
        this.events.set(eventId, updated);
        return { status: 200, body: updated };
      }
      case 'DELETE':
        this.events.delete(eventId);
        return { status: 204 };
      default:
        return this.error(405, 'MethodNotAllowed', `${method} is not supported`);
    }
  }

  private create(payload: Partial<Event>): Answer {
    const transactionId = payload.transactionId ?? undefined;
    if (transactionId && this.transactionIds.has(transactionId)) {
      const holder = this.duplicatesWithoutHolder ? null : this.transactionIds.get(transactionId)!;
      return { status: 400, body: { error: duplicateTransactionIdBody(holder) } };
    }

    const id = this.nextId();
    const created = { ...payload, id, changeKey: `ck-${id}-0` } as Event;
    this.events.set(id, created);
    if (transactionId) {
      this.transactionIds.set(transactionId, id);
    }

    if (this.lostCreateAnswers > 0) {
      this.lostCreateAnswers--;
      return this.error(502, 'UnknownError', '');
    }
    return { status: 201, body: created };
  }

  private calendarView(params: URLSearchParams): Answer {
    const from = toUtcMs(params.get('startDateTime'));
    const to = toUtcMs(params.get('endDateTime'));
    const top = Number(params.get('$top') ?? '10');
    const skip = Number(params.get('$skip') ?? '0');

    const inWindow = [...this.events.values()]
      .filter((e) => toUtcMs(e.start?.dateTime) < to && toUtcMs(e.end?.dateTime) > from)
      .sort((a, b) => toUtcMs(a.start?.dateTime) - toUtcMs(b.start?.dateTime));
    const page = inWindow.slice(skip, skip + top);

    const body: Record<string, unknown> = { value: page };
    if (skip + top < inWindow.length) {
      params.set('$skip', String(skip + top));
      body['@odata.nextLink'] = `${GRAPH_ROOT}/me/calendarView?${params.toString()}`;
    }
    return { status: 200, body };
  }

  private error(status: number, code: string, message: string): Answer {
    return { status, body: { error: { code, message } } };
  }

  private nextId(): string {
    this.seq += 1;
    return `AAkALgAAAAAAHYQDEapmEc2byACqAC-EWg0A-evt-${this.seq}`;
  }
}
