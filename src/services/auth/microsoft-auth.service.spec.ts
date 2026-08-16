import IoRedisMock from "ioredis-mock";
import axios from "axios";
import { EventEmitter2 } from "@nestjs/event-emitter";
import {
  MicrosoftAuthService,
  OutlookEmailBackfillResult,
} from "./microsoft-auth.service";
import { MicrosoftUser } from "../../entities/microsoft-user.entity";
import { MicrosoftUserStatus } from "../../enums/microsoft-user-status.enum";
import { OutlookEventTypes } from "../../enums/event-types.enum";
import {
  InMemoryOutlookLockStore,
  OutlookLockStore,
  RedisOutlookLockStore,
} from "../shared/outlook-lock.store";

/**
 * Regression coverage for ClickUp 86ca37pux — "multiple revocation emails sent
 * for some users". A burst of concurrent webhooks for one user each refreshes a
 * revoked token, each marks the user CORRUPTED, and (before the fix) each emitted
 * USER_REFRESH_TOKEN_INVALID, so the host app sent N revocation emails.
 *
 * The dedupe gate is the real OutlookLockStore, asserted against BOTH backends:
 *   - in-memory  → dedupe within a single process (multiple webhook events)
 *   - ioredis-mock → dedupe across "instances" (the ECS-fleet case)
 *
 * We exercise the real private markUserAsCorrupted / saveMicrosoftUser logic with
 * lightweight fakes for the unrelated collaborators, so the test runs the actual
 * concurrency gate rather than a mock of it.
 */

// markUserAsCorrupted / saveMicrosoftUser are private; expose just the two
// methods under test via a structural view (cast through unknown).
interface AnyService {
  markUserAsCorrupted(user: MicrosoftUser, reason: string): Promise<void>;
  saveMicrosoftUser(
    externalUserId: string,
    accessToken: string,
    refreshToken: string,
    expiresIn: number,
    scopes: string,
    outlookEmail: string | null,
  ): Promise<void>;
  fetchOutlookEmail(accessToken: string, correlationId: string): Promise<string | null>;
}

jest.mock("axios");
const mockedAxios = axios as jest.Mocked<typeof axios>;

const backends: Array<[string, () => Promise<OutlookLockStore>]> = [
  ["in-memory", async () => new InMemoryOutlookLockStore()],
  [
    "redis (ioredis-mock)",
    async () => {
      const client = new IoRedisMock();
      await client.flushall();
      return new RedisOutlookLockStore(client as never, "outlook:");
    },
  ],
];

const baseConfig = {
  clientId: "client-id",
  clientSecret: "client-secret",
  redirectPath: "http://localhost/callback",
  backendBaseUrl: "http://localhost",
};

function makeUser(overrides: Partial<MicrosoftUser> = {}): MicrosoftUser {
  const u = new MicrosoftUser();
  u.id = 1;
  u.externalUserId = "ext-1";
  u.status = MicrosoftUserStatus.ACTIVE;
  return Object.assign(u, overrides);
}

describe.each(backends)(
  "MicrosoftAuthService revocation-email dedupe (%s lock store)",
  (_name, makeStore) => {
    let service: AnyService;
    let lockStore: OutlookLockStore;
    let emit: jest.SpyInstance;
    let savedUser: MicrosoftUser;

    beforeEach(async () => {
      lockStore = await makeStore();
      const eventEmitter = new EventEmitter2();
      emit = jest.spyOn(eventEmitter, "emit");

      savedUser = makeUser();

      // Repository fake: save mutates an in-memory row; findOne returns it.
      const repo = {
        save: jest.fn(async (u: MicrosoftUser) => {
          savedUser = Object.assign(savedUser, u);
          return savedUser;
        }),
        update: jest.fn(async (_criteria: unknown, patch: Partial<MicrosoftUser>) => {
          savedUser = Object.assign(savedUser, patch);
          return { affected: 1 } as never;
        }),
        findOne: jest.fn(async () => savedUser),
      };

      service = new MicrosoftAuthService(
        eventEmitter,
        {} as never, // EmailService (unused on these paths)
        {} as never, // MicrosoftSubscriptionService (unused on these paths)
        baseConfig as never,
        {} as never, // csrfTokenRepository (unused on these paths)
        repo as never,
        lockStore,
      ) as unknown as AnyService;
    });

    it("emits USER_REFRESH_TOKEN_INVALID exactly once for a burst of concurrent corruptions", async () => {
      const user = makeUser();

      // 10 webhooks hit markUserAsCorrupted concurrently for the same user.
      await Promise.all(
        Array.from({ length: 10 }, () =>
          service.markUserAsCorrupted(user, "invalid_grant"),
        ),
      );

      const revocationEmits = emit.mock.calls.filter(
        (c) => c[0] === OutlookEventTypes.USER_REFRESH_TOKEN_INVALID,
      );
      expect(revocationEmits).toHaveLength(1);
      expect(revocationEmits[0][1]).toBe(user.externalUserId);
    });

    it("does not re-emit on a later corruption within the same cycle", async () => {
      const user = makeUser();

      await service.markUserAsCorrupted(user, "invalid_grant");
      await service.markUserAsCorrupted(user, "invalid_grant");

      const revocationEmits = emit.mock.calls.filter(
        (c) => c[0] === OutlookEventTypes.USER_REFRESH_TOKEN_INVALID,
      );
      expect(revocationEmits).toHaveLength(1);
    });

    it("emits again after re-auth resets the cycle (corrupt → re-auth → corrupt)", async () => {
      const user = makeUser();

      await service.markUserAsCorrupted(user, "invalid_grant");
      // Successful re-auth flips status to ACTIVE and clears the emit flag.
      await service.saveMicrosoftUser(
        user.externalUserId,
        "access",
        "refresh",
        3600,
        "scope",
        null,
      );
      await service.markUserAsCorrupted(user, "invalid_grant");

      const revocationEmits = emit.mock.calls.filter(
        (c) => c[0] === OutlookEventTypes.USER_REFRESH_TOKEN_INVALID,
      );
      expect(revocationEmits).toHaveLength(2);
    });

    it("default TTL (one week): repeated corruptions within the window never re-emit without re-auth", async () => {
      const user = makeUser();

      // baseConfig sets no revocationEmitFlagTtlMs → default one-week flag; well
      // within that window a second corruption is suppressed.
      await service.markUserAsCorrupted(user, "invalid_grant");
      await new Promise((r) => setTimeout(r, 40));
      await service.markUserAsCorrupted(user, "invalid_grant");

      const revocationEmits = emit.mock.calls.filter(
        (c) => c[0] === OutlookEventTypes.USER_REFRESH_TOKEN_INVALID,
      );
      expect(revocationEmits).toHaveLength(1);
    });

    it("re-emits after a configured finite TTL elapses (self-heal safety net)", async () => {
      const user = makeUser();
      const eventEmitter = new EventEmitter2();
      const ttlEmit = jest.spyOn(eventEmitter, "emit");
      const repo = {
        save: jest.fn(async (u: MicrosoftUser) => u),
        update: jest.fn(async () => ({ affected: 1 }) as never),
        findOne: jest.fn(async () => user),
      };
      const ttlService = new MicrosoftAuthService(
        eventEmitter,
        {} as never,
        {} as never,
        { ...baseConfig, revocationEmitFlagTtlMs: 30 } as never,
        {} as never,
        repo as never,
        lockStore,
      ) as unknown as AnyService;

      await ttlService.markUserAsCorrupted(user, "invalid_grant");
      await new Promise((r) => setTimeout(r, 60)); // outlast the 30ms TTL
      await ttlService.markUserAsCorrupted(user, "invalid_grant");

      const revocationEmits = ttlEmit.mock.calls.filter(
        (c) => c[0] === OutlookEventTypes.USER_REFRESH_TOKEN_INVALID,
      );
      expect(revocationEmits).toHaveLength(2);
    });

    it("does not emit when the DB save fails (flag is only set after a successful write)", async () => {
      const user = makeUser();
      // Force the next save to throw so markUserAsCorrupted returns early.
      const failingEmitter = new EventEmitter2();
      const failEmit = jest.spyOn(failingEmitter, "emit");
      const failingRepo = {
        save: jest.fn(async () => {
          throw new Error("db down");
        }),
        update: jest.fn(async () => {
          throw new Error("db down");
        }),
        findOne: jest.fn(async () => user),
      };
      const failingService = new MicrosoftAuthService(
        failingEmitter,
        {} as never,
        {} as never,
        baseConfig as never,
        {} as never,
        failingRepo as never,
        lockStore,
      ) as unknown as AnyService;

      await failingService.markUserAsCorrupted(user, "invalid_grant");

      const revocationEmits = failEmit.mock.calls.filter(
        (c) => c[0] === OutlookEventTypes.USER_REFRESH_TOKEN_INVALID,
      );
      expect(revocationEmits).toHaveLength(0);
    });
  },
);

/**
 * Regression coverage for "Outlook - 500 Error on calendar
 * disconnection". Legacy delegated rows created before the tokenExpiry column
 * existed still hold valid access/refresh tokens but have tokenExpiry = null.
 * The old guard rejected them as "app-only user or not yet authenticated",
 * turning every token fetch (webhook create/delete, reconciliation, sync) into a
 * failure — surfacing as an HTTP 500 on disconnect. A null expiry must instead be
 * treated as "unknown → refresh", not "no delegated credentials".
 */
interface TokenService {
  getUserAccessToken(params: {
    internalUserId?: number;
    externalUserId?: string;
    includeInactive?: boolean;
    cache?: boolean;
  }): Promise<string>;
  refreshAccessToken(refreshToken: string, internalUserId: number): Promise<string>;
  isTokenExpired(tokenExpiry: Date | null | undefined, bufferMinutes?: number): boolean;
}

describe("MicrosoftAuthService delegated-token guard (null tokenExpiry)", () => {
  function buildService(user: MicrosoftUser | null): TokenService {
    const repo = { save: jest.fn(), findOne: jest.fn(async () => user) };
    return new MicrosoftAuthService(
      new EventEmitter2(),
      {} as never, // EmailService
      {} as never, // MicrosoftSubscriptionService
      baseConfig as never,
      {} as never, // csrfTokenRepository
      repo as never,
      new InMemoryOutlookLockStore(),
    ) as unknown as TokenService;
  }

  it("isTokenExpired treats a null/undefined expiry as expired", () => {
    const service = buildService(null);
    expect(service.isTokenExpired(null)).toBe(true);
    expect(service.isTokenExpired(undefined)).toBe(true);
    // A comfortably-future expiry is still considered valid.
    expect(service.isTokenExpired(new Date(Date.now() + 60 * 60 * 1000))).toBe(false);
  });

  it("refreshes a legacy delegated user with null tokenExpiry instead of rejecting it", async () => {
    const user = makeUser({
      accessToken: "legacy-access",
      refreshToken: "legacy-refresh",
      tokenExpiry: null,
      scopes: "offline_access Calendars.ReadWrite",
    });
    const service = buildService(user);
    jest.spyOn(service, "refreshAccessToken").mockResolvedValue("refreshed-access");

    await expect(
      service.getUserAccessToken({ internalUserId: user.id }),
    ).resolves.toBe("refreshed-access");
    expect(service.refreshAccessToken).toHaveBeenCalledWith("legacy-refresh", user.id);
  });

  it("still rejects a genuine app-only user with no delegated tokens", async () => {
    const user = makeUser({ accessToken: null, refreshToken: null, tokenExpiry: null });
    const service = buildService(user);

    await expect(
      service.getUserAccessToken({ internalUserId: user.id }),
    ).rejects.toThrow(/has no delegated auth tokens/);
  });
});

/**
 * Coverage for capturing the connected Outlook mailbox email during the delegated
 * OAuth connect flow. The email is read from Graph /me (mail, falling back to
 * userPrincipalName), persisted on the microsoft_users row, and must be fail-open so a
 * transient /me failure never blocks a valid connection.
 */
interface EmailCaptureService {
  saveMicrosoftUser(
    externalUserId: string,
    accessToken: string,
    refreshToken: string,
    expiresIn: number,
    scopes: string,
    outlookEmail: string | null,
  ): Promise<void>;
  fetchOutlookEmail(accessToken: string, correlationId: string): Promise<string | null>;
}

describe("MicrosoftAuthService outlook email capture", () => {
  let savedUser: MicrosoftUser;
  let service: EmailCaptureService;

  beforeEach(() => {
    mockedAxios.get.mockReset();
    savedUser = makeUser();
    const repo = {
      save: jest.fn(async (u: MicrosoftUser) => {
        savedUser = Object.assign(savedUser, u);
        return savedUser;
      }),
      findOne: jest.fn(async () => null), // force the "create new user" branch
    };
    service = new MicrosoftAuthService(
      new EventEmitter2(),
      {} as never,
      {} as never,
      baseConfig as never,
      {} as never,
      repo as never,
      new InMemoryOutlookLockStore(),
    ) as unknown as EmailCaptureService;
  });

  it("fetchOutlookEmail returns mail from /me when present", async () => {
    mockedAxios.get.mockResolvedValue({
      data: { mail: "user@contoso.com", userPrincipalName: "user@contoso.onmicrosoft.com" },
    });
    await expect(service.fetchOutlookEmail("token", "corr-1")).resolves.toBe(
      "user@contoso.com",
    );
  });

  it("fetchOutlookEmail falls back to userPrincipalName when mail is null", async () => {
    mockedAxios.get.mockResolvedValue({
      data: { mail: null, userPrincipalName: "user@contoso.onmicrosoft.com" },
    });
    await expect(service.fetchOutlookEmail("token", "corr-2")).resolves.toBe(
      "user@contoso.onmicrosoft.com",
    );
  });

  it("fetchOutlookEmail is fail-open — returns null when /me rejects", async () => {
    mockedAxios.get.mockRejectedValue(new Error("graph down"));
    await expect(service.fetchOutlookEmail("token", "corr-3")).resolves.toBeNull();
  });

  it("saveMicrosoftUser persists the outlook email on the user row", async () => {
    await service.saveMicrosoftUser("ext-1", "access", "refresh", 3600, "scope", "user@contoso.com");
    expect(savedUser.outlookEmail).toBe("user@contoso.com");
  });

  it("saveMicrosoftUser persists null when no email was resolved (fail-open path)", async () => {
    await service.saveMicrosoftUser("ext-1", "access", "refresh", 3600, "scope", null);
    expect(savedUser.outlookEmail).toBeNull();
  });
});

/**
 * Coverage for backfillOutlookEmails — the one-off job that populates outlook_email for both
 * delegated and app-only users by exploiting their stored credentials. Asserts it routes token
 * resolution and the Graph endpoint by auth mode, updates only the email column, is
 * fault-tolerant per user, and terminates via its id cursor.
 */
interface BackfillService {
  backfillOutlookEmails(options?: {
    batchSize?: number;
    includeInactive?: boolean;
    maxUsers?: number;
    delayMsBetweenUsers?: number;
  }): Promise<OutlookEmailBackfillResult>;
}

describe("MicrosoftAuthService backfillOutlookEmails", () => {
  function makeQb(pages: MicrosoftUser[][]) {
    let call = 0;
    const qb: Record<string, unknown> = {
      leftJoinAndSelect: () => qb,
      where: () => qb,
      andWhere: () => qb,
      orderBy: () => qb,
      take: () => qb,
      getMany: async () => pages[call++] ?? [],
    };
    return qb;
  }

  // Fake MicrosoftSubscriptionService exposing only the auth-mode-aware resolver the backfill uses.
  function makeSubscription(
    resolve: (userId: number) => Promise<string | null>,
  ) {
    return {
      resolveUserAccessToken: jest.fn(
        async ({ internalUserId }: { internalUserId: number }) => resolve(internalUserId),
      ),
    };
  }

  function buildService(
    repo: Record<string, unknown>,
    subscription: { resolveUserAccessToken: jest.Mock },
  ): BackfillService {
    return new MicrosoftAuthService(
      new EventEmitter2(),
      {} as never, // EmailService
      subscription as never, // MicrosoftSubscriptionService
      baseConfig as never,
      {} as never, // csrfTokenRepository
      repo as never,
      new InMemoryOutlookLockStore(),
    ) as unknown as BackfillService;
  }

  it("delegated: updates via /me, skips empty profiles, counts per-user failures", async () => {
    const u1 = makeUser({ id: 1, externalUserId: "ext-1" }); // resolves a mail
    const u2 = makeUser({ id: 2, externalUserId: "ext-2" }); // /me returns nothing → skipped
    const u3 = makeUser({ id: 3, externalUserId: "ext-3" }); // token resolution throws → failed

    const qb = makeQb([[u1, u2, u3], []]);
    const update = jest.fn(async () => ({ affected: 1 }));
    const repo = { createQueryBuilder: jest.fn(() => qb), update };

    // u3's refresh token is invalid → resolver throws → counted as failed.
    const subscription = makeSubscription(async (userId) => {
      if (userId === 3) throw new Error("invalid_grant");
      return "delegated-token";
    });

    const service = buildService(repo, subscription);

    mockedAxios.get.mockReset();
    mockedAxios.get
      .mockResolvedValueOnce({ data: { mail: "one@contoso.com" } })
      .mockResolvedValueOnce({ data: { mail: null, userPrincipalName: null } });

    const result = await service.backfillOutlookEmails();

    expect(result).toEqual({ processed: 3, updated: 1, skipped: 1, failed: 1 });
    expect(update).toHaveBeenCalledTimes(1);
    expect(update).toHaveBeenCalledWith(1, { outlookEmail: "one@contoso.com" });
    // Delegated users are read from /me.
    expect(mockedAxios.get).toHaveBeenNthCalledWith(
      1,
      "https://graph.microsoft.com/v1.0/me",
      expect.anything(),
    );
  });

  it("app-only: resolves a tenant token and reads the profile from /users/{id}", async () => {
    const tenantUser = makeUser({
      id: 7,
      externalUserId: "ext-7",
      tenant: { tenantId: "tid-1" } as never,
      microsoftUserId: "ms-oid-xyz",
    });

    const qb = makeQb([[tenantUser], []]);
    const update = jest.fn(async () => ({ affected: 1 }));
    const repo = { createQueryBuilder: jest.fn(() => qb), update };
    const subscription = makeSubscription(async () => "app-only-token");
    const service = buildService(repo, subscription);

    mockedAxios.get.mockReset();
    mockedAxios.get.mockResolvedValueOnce({ data: { mail: "tenant@contoso.com" } });

    const result = await service.backfillOutlookEmails();

    expect(result).toEqual({ processed: 1, updated: 1, skipped: 0, failed: 0 });
    expect(update).toHaveBeenCalledWith(7, { outlookEmail: "tenant@contoso.com" });
    // App-only users can't use /me — the profile is read from /users/{microsoftUserId}.
    expect(mockedAxios.get).toHaveBeenCalledWith(
      "https://graph.microsoft.com/v1.0/users/ms-oid-xyz",
      expect.anything(),
    );
  });

  it("skips a user when the resolver returns no token (app-only not configured)", async () => {
    const tenantUser = makeUser({
      id: 8,
      externalUserId: "ext-8",
      tenant: { tenantId: "tid-2" } as never,
      microsoftUserId: "ms-oid-8",
    });
    const qb = makeQb([[tenantUser], []]);
    const update = jest.fn();
    const repo = { createQueryBuilder: jest.fn(() => qb), update };
    const subscription = makeSubscription(async () => null); // no token → inconclusive
    const service = buildService(repo, subscription);

    mockedAxios.get.mockReset();
    const result = await service.backfillOutlookEmails();

    expect(result).toEqual({ processed: 1, updated: 0, skipped: 1, failed: 0 });
    expect(update).not.toHaveBeenCalled();
    expect(mockedAxios.get).not.toHaveBeenCalled();
  });

  it("stops at the maxUsers cap and leaves the rest for a later run", async () => {
    const users = [1, 2, 3].map((id) => makeUser({ id, externalUserId: `ext-${id}` }));
    const qb = makeQb([users, []]);
    const update = jest.fn(async () => ({ affected: 1 }));
    const repo = { createQueryBuilder: jest.fn(() => qb), update };
    const subscription = makeSubscription(async () => "token");
    const service = buildService(repo, subscription);

    mockedAxios.get.mockReset();
    mockedAxios.get.mockResolvedValue({ data: { mail: "a@contoso.com" } });

    const result = await service.backfillOutlookEmails({ maxUsers: 1 });

    // Only the first user is touched; the resolver/Graph are not called for the rest.
    expect(result).toEqual({ processed: 1, updated: 1, skipped: 0, failed: 0 });
    expect(subscription.resolveUserAccessToken).toHaveBeenCalledTimes(1);
    expect(update).toHaveBeenCalledTimes(1);
  });

  it("terminates (does not loop forever) if the id cursor stops advancing", async () => {
    // A broken query that returns the SAME non-empty page on every fetch would loop forever
    // without the monotonic-progress backstop. getMany here never returns [].
    const stuckPage = [makeUser({ id: 1, externalUserId: "ext-1" })];
    const qb: Record<string, unknown> = {
      leftJoinAndSelect: () => qb,
      where: () => qb,
      andWhere: () => qb,
      orderBy: () => qb,
      take: () => qb,
      getMany: async () => stuckPage,
    };
    const update = jest.fn(async () => ({ affected: 1 }));
    const repo = { createQueryBuilder: jest.fn(() => qb), update };
    const subscription = makeSubscription(async () => "token");
    const service = buildService(repo, subscription);

    mockedAxios.get.mockReset();
    mockedAxios.get.mockResolvedValue({ data: { mail: "a@contoso.com" } });

    // If the backstop is missing this never resolves; the test would hang and fail on timeout.
    const result = await service.backfillOutlookEmails();

    // First page processed once, then the cursor can't advance → abort. No re-processing.
    expect(result.processed).toBe(1);
    expect(update).toHaveBeenCalledTimes(1);
  });

  it("returns a zero summary when there are no candidates", async () => {
    const qb = makeQb([[]]);
    const update = jest.fn();
    const repo = { createQueryBuilder: jest.fn(() => qb), update };
    const subscription = makeSubscription(async () => "token");
    const service = buildService(repo, subscription);

    const result = await service.backfillOutlookEmails();
    expect(result).toEqual({ processed: 0, updated: 0, skipped: 0, failed: 0 });
    expect(update).not.toHaveBeenCalled();
  });
});
