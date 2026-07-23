import { EventEmitter2 } from "@nestjs/event-emitter";
import { Response } from "express";
import { MicrosoftAuthController } from "./microsoft-auth.controller";
import { MicrosoftAuthService } from "../services/auth/microsoft-auth.service";
import { InMemoryOutlookLockStore } from "../services/shared/outlook-lock.store";

/**
 * End-to-end regression coverage for the Outlook OAuth state parameter:
 * use base64url for state + return 400 (not 500) on malformed state.
 *
 * Two bugs, both exercised through the REAL controller + REAL MicrosoftAuthService
 * (only the DB/lock collaborators are lightweight fakes, and the paths under test
 * never reach them):
 *
 *  1. Encoding: state was standard-base64 encoded and appended to the authorize URL
 *     raw. A '+' in the value decodes to a space when the callback query string is
 *     parsed → corrupted JSON → parseState fails for legitimate users. The fix uses
 *     URL-safe base64url end-to-end, so the value survives a query-string round-trip.
 *
 *  2. Error handling: a malformed/truncated state produced an uncaught throw that the
 *     controller answered with HTTP 500. Bot/scanner traffic replaying chopped callback
 *     URLs therefore paged as server errors. The fix returns a clean 400.
 */

function makeRes(): jest.Mocked<Response> {
  const res = {
    set: jest.fn().mockReturnThis(),
    send: jest.fn().mockReturnThis(),
    status: jest.fn().mockReturnThis(),
    json: jest.fn().mockReturnThis(),
  };
  return res as unknown as jest.Mocked<Response>;
}

const baseConfig = {
  clientId: "client-id",
  clientSecret: "client-secret",
  redirectPath: "/api/v1/auth/microsoft/callback",
  backendBaseUrl: "http://localhost",
};

function buildSut() {
  const eventEmitter = new EventEmitter2();

  // CSRF repo: getLoginUrl persists a token before returning; record it so the
  // round-trip assertion can confirm the state actually carried a csrf value.
  const csrfRepo = {
    saveToken: jest.fn(async () => undefined),
    findAndValidateToken: jest.fn(),
    cleanupExpiredTokens: jest.fn(),
  };

  const service = new MicrosoftAuthService(
    eventEmitter,
    {} as never, // EmailService (unused on these paths)
    {} as never, // MicrosoftSubscriptionService (unused on these paths)
    baseConfig as never,
    csrfRepo as never,
    {} as never, // MicrosoftUser repository (unused on these paths)
    new InMemoryOutlookLockStore(),
  );

  const controller = new MicrosoftAuthController(service);
  return { controller, service };
}

describe("MicrosoftAuth OAuth state (e2e)", () => {
  describe("encoding round-trip is URL-safe", () => {
    it("recovers the state after the value is carried in the callback query string", async () => {
      const { service } = buildSut();

      // "aaa>" is engineered to emit a '+' under STANDARD base64 — the exact class
      // of payload that used to corrupt when the query string decoded '+' to a space.
      const externalUserId = "aaa>";
      const authorizeUrl = await service.getLoginUrl(externalUserId);

      // Read `state` back exactly as the callback (Express) would: URLSearchParams
      // applies percent + '+'-to-space decoding.
      const stateFromCallback = new URL(authorizeUrl).searchParams.get("state");
      expect(stateFromCallback).toBeTruthy();

      const parsed = service.parseState(stateFromCallback as string);
      expect(parsed).not.toBeNull();
      expect(parsed?.userId).toBe(externalUserId);
      expect(parsed?.csrf).toEqual(expect.any(String));
    });

    it("emits only URL-safe base64url characters (no '+', '/', or '=')", async () => {
      const { service } = buildSut();

      const authorizeUrl = await service.getLoginUrl("3646");
      // The raw, still-percent-encoded value straight out of the query string.
      const rawState = new URLSearchParams(authorizeUrl.split("?")[1]).get("state");

      expect(rawState).toMatch(/^[A-Za-z0-9_-]+$/);
    });
  });

  describe("malformed state returns 400, not 500", () => {
    it("answers 400 for a truncated state payload (the prod-incident shape)", async () => {
      const { controller } = buildSut();
      const res = makeRes();

      // A base64url value that decodes to TRUNCATED JSON — mirrors the real
      // incident where the callback URL was chopped mid-csrf.
      const truncated = Buffer.from(
        '{"userId":"3646","csrf":"8e71c558213c6e812a54d75674',
      ).toString("base64url");

      await controller.handleOauthCallback("any-code", truncated, res);

      expect(res.status).toHaveBeenCalledWith(400);
      expect(res.status).not.toHaveBeenCalledWith(500);
    });

    it("answers 400 for a garbage (non-JSON) state", async () => {
      const { controller } = buildSut();
      const res = makeRes();

      await controller.handleOauthCallback("any-code", "!!!not-base64!!!", res);

      expect(res.status).toHaveBeenCalledWith(400);
      expect(res.status).not.toHaveBeenCalledWith(500);
    });

    it("answers 400 when state is missing entirely", async () => {
      const { controller } = buildSut();
      const res = makeRes();

      await controller.handleOauthCallback("any-code", undefined as never, res);

      expect(res.status).toHaveBeenCalledWith(400);
    });
  });
});
