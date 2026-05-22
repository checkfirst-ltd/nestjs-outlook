import { DeltaSyncService, DeltaItem } from "./delta-sync.service";
import { OutlookDeltaLinkRepository } from "../../repositories/outlook-delta-link.repository";
import { UserIdConverterService } from "./user-id-converter.service";
import { GraphRateLimiterService } from "./graph-rate-limiter.service";
import { EventEmitter2 } from "@nestjs/event-emitter";
import { Client } from "@microsoft/microsoft-graph-client";
import { ResourceType } from "../../enums/resource-type.enum";

type Page = { items: DeltaItem[]; deltaLink: string | null; isLastPage: boolean };

function makePageGenerator(pages: Page[]) {
  return async function* () {
    for (const page of pages) {
      yield page;
    }
  };
}

/**
 * Returns an async-generator factory that throws `error` on first iteration,
 * without yielding any pages. Used to simulate Graph SDK errors from
 * fetchDeltaPagesCore.
 */
function makeThrowingPageGenerator(error: unknown) {
  return async function* (): AsyncGenerator<Page, void, unknown> {
    await Promise.resolve();
    throw error;
  };
}

/**
 * Builds an Error shaped like the Microsoft Graph SDK's 410 Gone response
 * (statusCode at the top level), which is what `is410Error` looks for.
 */
function make410Error(): Error & { statusCode: number } {
  const err = new Error("Sync state not found") as Error & { statusCode: number };
  err.statusCode = 410;
  return err;
}

describe("DeltaSyncService.streamDeltaChanges — skipCursorAdvanceOnEmpty guard", () => {
  let service: DeltaSyncService;
  let deltaLinkRepository: jest.Mocked<OutlookDeltaLinkRepository>;
  let userIdConverter: jest.Mocked<UserIdConverterService>;
  let rateLimiter: jest.Mocked<GraphRateLimiterService>;
  let eventEmitter: jest.Mocked<EventEmitter2>;
  let client: Client;

  const EXTERNAL_USER_ID = "3617";
  const INTERNAL_USER_ID = 875;
  const REQUEST_URL = "/me/events/delta";
  const OLD_DELTA_LINK = "https://graph.microsoft.com/v1.0/me/events/delta?$deltatoken=OLD";
  const NEW_DELTA_LINK = "https://graph.microsoft.com/v1.0/me/events/delta?$deltatoken=NEW";

  beforeEach(() => {
    deltaLinkRepository = {
      getDeltaLink: jest.fn(),
      saveDeltaLink: jest.fn().mockResolvedValue(undefined),
      deleteDeltaLink: jest.fn().mockResolvedValue(undefined),
    } as unknown as jest.Mocked<OutlookDeltaLinkRepository>;

    userIdConverter = {
      externalToInternal: jest.fn().mockResolvedValue(INTERNAL_USER_ID),
    } as unknown as jest.Mocked<UserIdConverterService>;

    rateLimiter = {
      acquirePermit: jest.fn().mockResolvedValue(undefined),
    } as unknown as jest.Mocked<GraphRateLimiterService>;

    eventEmitter = {
      emit: jest.fn(),
    } as unknown as jest.Mocked<EventEmitter2>;

    client = {} as Client;

    service = new DeltaSyncService(
      deltaLinkRepository,
      userIdConverter,
      rateLimiter,
      eventEmitter,
    );
  });

  /**
   * Drains an async generator. Returns all yielded items in a flat array plus the return value.
   */
  async function drain<T, TReturn>(
    gen: AsyncGenerator<T[], TReturn, unknown>,
  ): Promise<{ items: T[]; returnValue: TReturn }> {
    const items: T[] = [];
    let next = await gen.next();
    while (!next.done) {
      items.push(...(next.value as T[]));
      next = await gen.next();
    }
    return { items, returnValue: next.value as TReturn };
  }

  it("(a) empty incremental result + skipCursorAdvanceOnEmpty=true → KEEPS old cursor", async () => {
    // Existing cursor in DB
    deltaLinkRepository.getDeltaLink.mockResolvedValue(OLD_DELTA_LINK);

    // Graph returns one empty page with a new deltaLink
    const pages: Page[] = [
      { items: [], deltaLink: NEW_DELTA_LINK, isLastPage: true },
    ];
    jest
      .spyOn(service as any, "fetchDeltaPagesCore")
      .mockImplementation(makePageGenerator(pages));

    const gen = service.streamDeltaChanges(
      client,
      REQUEST_URL,
      EXTERNAL_USER_ID,
      /* forceReset */ false,
      /* dateRange */ undefined,
      /* saveDeltaLink */ true,
      /* skipCursorAdvanceOnEmpty */ true,
    );

    const { items, returnValue } = await drain(gen);

    expect(items).toHaveLength(0);
    expect(returnValue).toBe(NEW_DELTA_LINK);
    // Guard fires: save was NOT called
    expect(deltaLinkRepository.saveDeltaLink).not.toHaveBeenCalled();
    // Old cursor untouched
    expect(deltaLinkRepository.deleteDeltaLink).not.toHaveBeenCalled();
  });

  it("(b) non-empty incremental result + skipCursorAdvanceOnEmpty=true → SAVES new cursor", async () => {
    deltaLinkRepository.getDeltaLink.mockResolvedValue(OLD_DELTA_LINK);

    const pages: Page[] = [
      {
        items: [{ id: "evt-1", lastModifiedDateTime: "2026-05-21T10:00:00Z" }],
        deltaLink: NEW_DELTA_LINK,
        isLastPage: true,
      },
    ];
    jest
      .spyOn(service as any, "fetchDeltaPagesCore")
      .mockImplementation(makePageGenerator(pages));

    const gen = service.streamDeltaChanges(
      client,
      REQUEST_URL,
      EXTERNAL_USER_ID,
      false,
      undefined,
      true,
      true,
    );

    const { items, returnValue } = await drain(gen);

    expect(items).toHaveLength(1);
    expect(returnValue).toBe(NEW_DELTA_LINK);
    expect(deltaLinkRepository.saveDeltaLink).toHaveBeenCalledTimes(1);
    expect(deltaLinkRepository.saveDeltaLink).toHaveBeenCalledWith(
      INTERNAL_USER_ID,
      ResourceType.CALENDAR,
      NEW_DELTA_LINK,
    );
  });

  it("(c) empty incremental result + skipCursorAdvanceOnEmpty=false (default) → SAVES new cursor (today's behavior preserved)", async () => {
    deltaLinkRepository.getDeltaLink.mockResolvedValue(OLD_DELTA_LINK);

    const pages: Page[] = [
      { items: [], deltaLink: NEW_DELTA_LINK, isLastPage: true },
    ];
    jest
      .spyOn(service as any, "fetchDeltaPagesCore")
      .mockImplementation(makePageGenerator(pages));

    const gen = service.streamDeltaChanges(
      client,
      REQUEST_URL,
      EXTERNAL_USER_ID,
      false,
      undefined,
      true,
      // skipCursorAdvanceOnEmpty omitted → defaults to false
    );

    const { items, returnValue } = await drain(gen);

    expect(items).toHaveLength(0);
    expect(returnValue).toBe(NEW_DELTA_LINK);
    // No guard: save was called as today
    expect(deltaLinkRepository.saveDeltaLink).toHaveBeenCalledTimes(1);
    expect(deltaLinkRepository.saveDeltaLink).toHaveBeenCalledWith(
      INTERNAL_USER_ID,
      ResourceType.CALENDAR,
      NEW_DELTA_LINK,
    );
  });

  it("(d) first-time init (no startLink) + empty result + skipCursorAdvanceOnEmpty=true → SAVES new cursor (baseline must be established)", async () => {
    // No existing cursor
    deltaLinkRepository.getDeltaLink.mockResolvedValue(null);

    const pages: Page[] = [
      { items: [], deltaLink: NEW_DELTA_LINK, isLastPage: true },
    ];
    jest
      .spyOn(service as any, "fetchDeltaPagesCore")
      .mockImplementation(makePageGenerator(pages));

    const gen = service.streamDeltaChanges(
      client,
      REQUEST_URL,
      EXTERNAL_USER_ID,
      false,
      undefined,
      true,
      true,
    );

    const { items, returnValue } = await drain(gen);

    expect(items).toHaveLength(0);
    expect(returnValue).toBe(NEW_DELTA_LINK);
    // Guard does NOT fire on first-time init: save IS called
    expect(deltaLinkRepository.saveDeltaLink).toHaveBeenCalledTimes(1);
    expect(deltaLinkRepository.saveDeltaLink).toHaveBeenCalledWith(
      INTERNAL_USER_ID,
      ResourceType.CALENDAR,
      NEW_DELTA_LINK,
    );
  });

  it("(e) forceReset path + empty result + skipCursorAdvanceOnEmpty=true → SAVES new cursor (reset by definition has no prior cursor to fall back to)", async () => {
    // Existing cursor will be deleted by forceReset
    deltaLinkRepository.getDeltaLink.mockResolvedValue(OLD_DELTA_LINK);

    const pages: Page[] = [
      { items: [], deltaLink: NEW_DELTA_LINK, isLastPage: true },
    ];
    jest
      .spyOn(service as any, "fetchDeltaPagesCore")
      .mockImplementation(makePageGenerator(pages));

    const gen = service.streamDeltaChanges(
      client,
      REQUEST_URL,
      EXTERNAL_USER_ID,
      /* forceReset */ true,
      undefined,
      true,
      true,
    );

    const { items, returnValue } = await drain(gen);

    expect(items).toHaveLength(0);
    expect(returnValue).toBe(NEW_DELTA_LINK);
    // forceReset deletes the old link first
    expect(deltaLinkRepository.deleteDeltaLink).toHaveBeenCalledTimes(1);
    // And the new link IS saved (guard skipped because !usedExistingCursor after reset)
    expect(deltaLinkRepository.saveDeltaLink).toHaveBeenCalledTimes(1);
    expect(deltaLinkRepository.saveDeltaLink).toHaveBeenCalledWith(
      INTERNAL_USER_ID,
      ResourceType.CALENDAR,
      NEW_DELTA_LINK,
    );
  });

  it("(f) 410 Gone recovery → deletes expired cursor, restarts with full sync, yields recovery items, saves recovery cursor, emits with isRecovery: true", async () => {
    const RECOVERY_DELTA_LINK = "https://graph.microsoft.com/v1.0/me/events/delta?$deltatoken=RECOVERY";

    deltaLinkRepository.getDeltaLink.mockResolvedValue(OLD_DELTA_LINK);

    // First call to fetchDeltaPagesCore: throws 410-shaped error before yielding anything.
    // Second call (the recovery): yields one normal page with new deltaLink.
    const recoveryPages: Page[] = [
      {
        items: [{ id: "evt-r1", lastModifiedDateTime: "2026-05-22T08:00:00Z" }],
        deltaLink: RECOVERY_DELTA_LINK,
        isLastPage: true,
      },
    ];

    const spy = jest.spyOn(service as any, "fetchDeltaPagesCore");
    spy.mockImplementationOnce(makeThrowingPageGenerator(make410Error()));
    spy.mockImplementationOnce(makePageGenerator(recoveryPages));

    const gen = service.streamDeltaChanges(
      client,
      REQUEST_URL,
      EXTERNAL_USER_ID,
      false,
      undefined,
      true,
      false,
    );

    const { items, returnValue } = await drain(gen);

    expect(items).toHaveLength(1);
    expect(items[0].id).toBe("evt-r1");
    expect(returnValue).toBe(RECOVERY_DELTA_LINK);

    // Recovery deleted the expired cursor exactly once
    expect(deltaLinkRepository.deleteDeltaLink).toHaveBeenCalledTimes(1);
    expect(deltaLinkRepository.deleteDeltaLink).toHaveBeenCalledWith(
      INTERNAL_USER_ID,
      ResourceType.CALENDAR,
    );

    // Recovery cursor was saved
    expect(deltaLinkRepository.saveDeltaLink).toHaveBeenCalledTimes(1);
    expect(deltaLinkRepository.saveDeltaLink).toHaveBeenCalledWith(
      INTERNAL_USER_ID,
      ResourceType.CALENDAR,
      RECOVERY_DELTA_LINK,
    );

    // Progress emit on the recovery path carries isRecovery: true
    expect(eventEmitter.emit).toHaveBeenCalledWith(
      "delta-sync.page-processed",
      expect.objectContaining({ isRecovery: true, itemsInPage: 1 }),
    );
  });

  it("(g) multi-page sync (skiptoken chain) → yields items from all pages in order, saves only the final cursor once", async () => {
    deltaLinkRepository.getDeltaLink.mockResolvedValue(OLD_DELTA_LINK);

    // Three pages: only the last carries deltaLink + isLastPage.
    // Items chosen so each page's earliest event is later than the previous page's latest
    // (delta is paginated, not strictly sorted across pages — sort applies within page).
    const pages: Page[] = [
      {
        items: [
          { id: "p1-a", lastModifiedDateTime: "2026-05-22T08:00:00Z" },
          { id: "p1-b", lastModifiedDateTime: "2026-05-22T08:05:00Z" },
        ],
        deltaLink: null,
        isLastPage: false,
      },
      {
        items: [
          { id: "p2-a", lastModifiedDateTime: "2026-05-22T09:00:00Z" },
        ],
        deltaLink: null,
        isLastPage: false,
      },
      {
        items: [
          { id: "p3-a", lastModifiedDateTime: "2026-05-22T10:00:00Z" },
          { id: "p3-b", lastModifiedDateTime: "2026-05-22T10:05:00Z" },
        ],
        deltaLink: NEW_DELTA_LINK,
        isLastPage: true,
      },
    ];
    jest
      .spyOn(service as any, "fetchDeltaPagesCore")
      .mockImplementation(makePageGenerator(pages));

    const gen = service.streamDeltaChanges(
      client,
      REQUEST_URL,
      EXTERNAL_USER_ID,
      false,
      undefined,
      true,
      false,
    );

    const { items, returnValue } = await drain(gen);

    expect(items.map((it) => it.id)).toEqual(["p1-a", "p1-b", "p2-a", "p3-a", "p3-b"]);
    expect(returnValue).toBe(NEW_DELTA_LINK);

    // Save called exactly once, with the deltaLink from the LAST page
    expect(deltaLinkRepository.saveDeltaLink).toHaveBeenCalledTimes(1);
    expect(deltaLinkRepository.saveDeltaLink).toHaveBeenCalledWith(
      INTERNAL_USER_ID,
      ResourceType.CALENDAR,
      NEW_DELTA_LINK,
    );

    // Three progress emits, none with isRecovery: true on the main path
    expect(eventEmitter.emit).toHaveBeenCalledTimes(3);
    for (const call of (eventEmitter.emit as jest.Mock).mock.calls) {
      expect(call[1]).not.toHaveProperty("isRecovery", true);
    }
  });

  it("(h) dateRange first-init → fetchDeltaPagesCore is called with a URL containing startDateTime and endDateTime", async () => {
    // No existing cursor → first-time init path
    deltaLinkRepository.getDeltaLink.mockResolvedValue(null);

    const startDate = new Date("2026-05-22T00:00:00Z");
    const endDate = new Date("2031-05-22T00:00:00Z");

    const pages: Page[] = [
      { items: [], deltaLink: NEW_DELTA_LINK, isLastPage: true },
    ];
    const spy = jest
      .spyOn(service as any, "fetchDeltaPagesCore")
      .mockImplementation(makePageGenerator(pages));

    const gen = service.streamDeltaChanges(
      client,
      REQUEST_URL,
      EXTERNAL_USER_ID,
      false,
      { startDate, endDate },
      true,
      false,
    );

    await drain(gen);

    expect(spy).toHaveBeenCalledTimes(1);
    const startUrlArg = spy.mock.calls[0][1] as string;
    expect(startUrlArg).toContain(REQUEST_URL);
    expect(startUrlArg).toContain(`startDateTime=${startDate.toISOString()}`);
    expect(startUrlArg).toContain(`endDateTime=${endDate.toISOString()}`);
  });

  it("(i) non-410 error propagates → no save, no delete, drain rejects with the original error", async () => {
    deltaLinkRepository.getDeltaLink.mockResolvedValue(OLD_DELTA_LINK);

    jest
      .spyOn(service as any, "fetchDeltaPagesCore")
      .mockImplementation(makeThrowingPageGenerator(new Error("boom")));

    const gen = service.streamDeltaChanges(
      client,
      REQUEST_URL,
      EXTERNAL_USER_ID,
      false,
      undefined,
      true,
      false,
    );

    await expect(drain(gen)).rejects.toThrow("boom");
    expect(deltaLinkRepository.saveDeltaLink).not.toHaveBeenCalled();
    expect(deltaLinkRepository.deleteDeltaLink).not.toHaveBeenCalled();
  });

  it("(j) saveDeltaLink=false → no save call even with non-empty incremental result, return value still equals the new deltaLink", async () => {
    deltaLinkRepository.getDeltaLink.mockResolvedValue(OLD_DELTA_LINK);

    const pages: Page[] = [
      {
        items: [{ id: "evt-1", lastModifiedDateTime: "2026-05-22T10:00:00Z" }],
        deltaLink: NEW_DELTA_LINK,
        isLastPage: true,
      },
    ];
    jest
      .spyOn(service as any, "fetchDeltaPagesCore")
      .mockImplementation(makePageGenerator(pages));

    const gen = service.streamDeltaChanges(
      client,
      REQUEST_URL,
      EXTERNAL_USER_ID,
      false,
      undefined,
      /* saveDeltaLink */ false,
      false,
    );

    const { items, returnValue } = await drain(gen);

    expect(items).toHaveLength(1);
    expect(returnValue).toBe(NEW_DELTA_LINK);
    expect(deltaLinkRepository.saveDeltaLink).not.toHaveBeenCalled();
  });

  it("(k) sortDeltaItems applied → items yielded oldest→newest within a page even if input is reverse-chronological", async () => {
    deltaLinkRepository.getDeltaLink.mockResolvedValue(OLD_DELTA_LINK);

    // Input order: newest first. After sort: oldest first.
    const pages: Page[] = [
      {
        items: [
          { id: "evt-newest", lastModifiedDateTime: "2026-05-22T12:00:00Z" },
          { id: "evt-middle", lastModifiedDateTime: "2026-05-22T10:00:00Z" },
          { id: "evt-oldest", lastModifiedDateTime: "2026-05-22T08:00:00Z" },
        ],
        deltaLink: NEW_DELTA_LINK,
        isLastPage: true,
      },
    ];
    jest
      .spyOn(service as any, "fetchDeltaPagesCore")
      .mockImplementation(makePageGenerator(pages));

    const gen = service.streamDeltaChanges(
      client,
      REQUEST_URL,
      EXTERNAL_USER_ID,
      false,
      undefined,
      true,
      false,
    );

    const { items } = await drain(gen);

    expect(items.map((it) => it.id)).toEqual(["evt-oldest", "evt-middle", "evt-newest"]);
  });
});
