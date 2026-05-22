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
});
