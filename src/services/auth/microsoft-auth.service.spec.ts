import IoRedisMock from "ioredis-mock";
import { EventEmitter2 } from "@nestjs/event-emitter";
import { MicrosoftAuthService } from "./microsoft-auth.service";
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
  ): Promise<void>;
}

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
