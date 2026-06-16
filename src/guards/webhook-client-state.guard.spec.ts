import { ExecutionContext } from "@nestjs/common";
import { EventEmitter2 } from "@nestjs/event-emitter";
import {
  WebhookClientStateGuard,
  RequestWithWebhookValidation,
  WebhookRejectedEvent,
} from "./webhook-client-state.guard";
import { OutlookWebhookSubscriptionRepository } from "../repositories/outlook-webhook-subscription.repository";
import { OutlookEventTypes } from "../enums/event-types.enum";

type FakeRequest = Partial<RequestWithWebhookValidation> & {
  query?: Record<string, unknown>;
  body?: unknown;
  path?: string;
};

function makeContext(request: FakeRequest): ExecutionContext {
  return {
    switchToHttp: () => ({ getRequest: () => request }),
  } as unknown as ExecutionContext;
}

describe("WebhookClientStateGuard", () => {
  let guard: WebhookClientStateGuard;
  let repo: { findBySubscriptionId: jest.Mock };
  let emitter: { emit: jest.Mock };

  beforeEach(() => {
    repo = { findBySubscriptionId: jest.fn() };
    emitter = { emit: jest.fn() };
    guard = new WebhookClientStateGuard(
      repo as unknown as OutlookWebhookSubscriptionRepository,
      emitter as unknown as EventEmitter2,
    );
  });

  function emittedRejections(): WebhookRejectedEvent[] {
    return emitter.emit.mock.calls
      .filter(([type]) => type === OutlookEventTypes.WEBHOOK_REJECTED)
      .map(([, payload]) => payload as WebhookRejectedEvent);
  }

  it("allows validation requests through without touching the repo or emitting", async () => {
    const req: FakeRequest = { query: { validationToken: "abc" }, body: {} };

    await expect(guard.canActivate(makeContext(req))).resolves.toBe(true);
    expect(repo.findBySubscriptionId).not.toHaveBeenCalled();
    expect(emitter.emit).not.toHaveBeenCalled();
  });

  it("allows an empty notification body through", async () => {
    const req: FakeRequest = { query: {}, body: { value: [] } };

    await expect(guard.canActivate(makeContext(req))).resolves.toBe(true);
    expect(emitter.emit).not.toHaveBeenCalled();
  });

  it("marks a valid clientState as valid and emits nothing", async () => {
    repo.findBySubscriptionId.mockResolvedValue({
      subscriptionId: "sub-1",
      clientState: "secret-state",
      userId: 42,
    });
    const req: FakeRequest = {
      query: {},
      path: "/calendar/webhook/notification",
      body: { value: [{ subscriptionId: "sub-1", clientState: "secret-state" }] },
    };

    await expect(guard.canActivate(makeContext(req))).resolves.toBe(true);
    expect(req.webhookValidation).toEqual({ valid: true, invalidItems: [] });
    expect(emittedRejections()).toHaveLength(0);
  });

  it("rejects (and emits) a missing subscriptionId", async () => {
    const req: FakeRequest = {
      query: {},
      path: "/calendar/webhook/notification",
      body: { value: [{ clientState: "whatever" }] },
    };

    await expect(guard.canActivate(makeContext(req))).resolves.toBe(true);
    expect(req.webhookValidation?.valid).toBe(false);
    expect(req.webhookValidation?.invalidItems[0]).toMatchObject({
      index: 0,
      reason: "missing_subscription_id",
    });
    const events = emittedRejections();
    expect(events).toHaveLength(1);
    expect(events[0]).toMatchObject({
      reason: "missing_subscription_id",
      subscriptionId: "unknown",
      userId: null,
      endpoint: "/calendar/webhook/notification",
    });
  });

  it("rejects (and emits) an unknown subscription", async () => {
    repo.findBySubscriptionId.mockResolvedValue(null);
    const req: FakeRequest = {
      query: {},
      path: "/calendar/webhook/notification",
      body: { value: [{ subscriptionId: "forged", clientState: "x" }] },
    };

    await expect(guard.canActivate(makeContext(req))).resolves.toBe(true);
    expect(req.webhookValidation?.valid).toBe(false);
    const events = emittedRejections();
    expect(events).toHaveLength(1);
    expect(events[0]).toMatchObject({
      reason: "unknown_subscription",
      subscriptionId: "forged",
      userId: null,
    });
  });

  it("rejects (and emits) a clientState mismatch", async () => {
    repo.findBySubscriptionId.mockResolvedValue({
      subscriptionId: "sub-1",
      clientState: "secret-state",
      userId: 42,
    });
    const req: FakeRequest = {
      query: {},
      path: "/calendar/webhook/notification",
      body: { value: [{ subscriptionId: "sub-1", clientState: "wrong-state" }] },
    };

    await expect(guard.canActivate(makeContext(req))).resolves.toBe(true);
    expect(req.webhookValidation?.valid).toBe(false);
    const events = emittedRejections();
    expect(events).toHaveLength(1);
    expect(events[0]).toMatchObject({
      reason: "client_state_mismatch",
      subscriptionId: "sub-1",
      userId: 42,
    });
  });

  it("isolates a forged item in a mixed batch (one valid, one mismatch)", async () => {
    repo.findBySubscriptionId.mockImplementation((id: string) =>
      Promise.resolve(
        id === "legit"
          ? { subscriptionId: "legit", clientState: "ok", userId: 7 }
          : { subscriptionId: "forged", clientState: "real", userId: 9 },
      ),
    );
    const req: FakeRequest = {
      query: {},
      path: "/calendar/webhook/notification",
      body: {
        value: [
          { subscriptionId: "legit", clientState: "ok" },
          { subscriptionId: "forged", clientState: "tampered" },
        ],
      },
    };

    await guard.canActivate(makeContext(req));

    expect(req.webhookValidation?.valid).toBe(false);
    expect(req.webhookValidation?.invalidItems).toHaveLength(1);
    expect(req.webhookValidation?.invalidItems[0]).toMatchObject({
      index: 1,
      reason: "client_state_mismatch",
    });
    expect(emittedRejections()).toHaveLength(1);
  });
});
