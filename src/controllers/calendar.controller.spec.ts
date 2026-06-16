import { Test } from "@nestjs/testing";
import { EventEmitter2 } from "@nestjs/event-emitter";
import { Response } from "express";
import { CalendarController } from "./calendar.controller";
import { CalendarService } from "../services/calendar/calendar.service";
import { LifecycleEventHandlerService } from "../services/calendar/lifecycle-event-handler.service";
import { OutlookWebhookNotificationDto } from "../dto/outlook-webhook-notification.dto";
import {
  WebhookClientStateGuard,
  WebhookValidationResult,
} from "../guards/webhook-client-state.guard";
import { OutlookWebhookSubscriptionRepository } from "../repositories/outlook-webhook-subscription.repository";

function makeRes(): jest.Mocked<Response> {
  const res = {
    set: jest.fn().mockReturnThis(),
    send: jest.fn().mockReturnThis(),
    status: jest.fn().mockReturnThis(),
    json: jest.fn().mockReturnThis(),
  };
  return res as unknown as jest.Mocked<Response>;
}

const req = {} as never;

/** Build a request carrying the guard's verdict, as the controller reads it. */
function reqWithValidation(validation: WebhookValidationResult) {
  return { webhookValidation: validation } as never;
}

function calendarItem(overrides: Record<string, unknown> = {}) {
  return {
    subscriptionId: "sub-1",
    subscriptionExpirationDateTime: "2030-01-01T00:00:00Z",
    changeType: "updated",
    resource: "/users/u/events/evt-1",
    resourceData: {
      "@odata.type": "#microsoft.graph.event",
      id: "evt-1",
    },
    clientState: "secret-state",
    tenantId: "tenant-1",
    ...overrides,
  };
}

describe("CalendarController (webhook receiving)", () => {
  let controller: CalendarController;
  let calendarService: {
    handleOutlookWebhookV2: jest.Mock;
    handleOutlookWebhook: jest.Mock;
  };
  let lifecycle: { handleLifecycleEvent: jest.Mock };

  beforeEach(async () => {
    calendarService = {
      handleOutlookWebhookV2: jest
        .fn()
        .mockResolvedValue({ success: true, message: "ok" }),
      handleOutlookWebhook: jest
        .fn()
        .mockResolvedValue({ success: true, message: "ok" }),
    };
    lifecycle = {
      handleLifecycleEvent: jest.fn().mockResolvedValue({ success: true }),
    };

    const moduleRef = await Test.createTestingModule({
      controllers: [CalendarController],
      providers: [
        { provide: CalendarService, useValue: calendarService },
        { provide: LifecycleEventHandlerService, useValue: lifecycle },
        {
          provide: OutlookWebhookSubscriptionRepository,
          useValue: { findBySubscriptionId: jest.fn() },
        },
        { provide: EventEmitter2, useValue: { emit: jest.fn() } },
        WebhookClientStateGuard,
      ],
    }).compile();

    controller = moduleRef.get(CalendarController);
  });

  describe("POST /calendar/webhook/notification", () => {
    it("echoes the decoded validation token as text/plain and does not process", async () => {
      const res = makeRes();
      await controller.handleCalendarWebhookNotification(
        "abc%20123",
        {} as OutlookWebhookNotificationDto,
        req,
        res,
      );

      expect(res.set).toHaveBeenCalledWith(
        "Content-Type",
        "text/plain; charset=utf-8",
      );
      expect(res.send).toHaveBeenCalledWith("abc 123");
      expect(calendarService.handleOutlookWebhookV2).not.toHaveBeenCalled();
    });

    it("processes a single notification via handleOutlookWebhookV2 with mapped fields", async () => {
      const res = makeRes();
      const body = {
        value: [calendarItem()],
      } as unknown as OutlookWebhookNotificationDto;

      await controller.handleCalendarWebhookNotification("", body, req, res);

      expect(calendarService.handleOutlookWebhookV2).toHaveBeenCalledTimes(1);
      const [mapped, traceId] =
        calendarService.handleOutlookWebhookV2.mock.calls[0];
      expect(mapped).toMatchObject({
        subscriptionId: "sub-1",
        changeType: "updated",
        resource: "/users/u/events/evt-1",
        clientState: "secret-state",
        tenantId: "tenant-1",
      });
      expect(mapped.resourceData).toMatchObject({ id: "evt-1" });
      expect(typeof traceId).toBe("string");
      expect(res.json).toHaveBeenCalledWith(
        expect.objectContaining({ success: true }),
      );
    });

    it("returns 202 and processes each item for a batch of >2", async () => {
      const res = makeRes();
      const body = {
        value: [
          calendarItem({ subscriptionId: "s1" }),
          calendarItem({ subscriptionId: "s2" }),
          calendarItem({ subscriptionId: "s3" }),
        ],
      } as unknown as OutlookWebhookNotificationDto;

      await controller.handleCalendarWebhookNotification("", body, req, res);

      expect(res.status).toHaveBeenCalledWith(202);
      expect(calendarService.handleOutlookWebhookV2).toHaveBeenCalledTimes(3);
    });

    it("returns 500 when processing throws", async () => {
      const res = makeRes();
      calendarService.handleOutlookWebhookV2.mockRejectedValueOnce(
        new Error("boom"),
      );
      const body = {
        value: [calendarItem()],
      } as unknown as OutlookWebhookNotificationDto;

      await controller.handleCalendarWebhookNotification("", body, req, res);

      expect(res.status).toHaveBeenCalledWith(500);
      expect(res.json).toHaveBeenCalledWith(
        expect.objectContaining({ success: false }),
      );
    });
  });

  describe("POST /calendar/webhook", () => {
    it("echoes the decoded validation token as text/plain", async () => {
      const res = makeRes();
      await controller.handleCalendarWebhook(
        "tok%2Fen",
        {} as OutlookWebhookNotificationDto,
        req,
        res,
      );

      expect(res.set).toHaveBeenCalledWith(
        "Content-Type",
        "text/plain; charset=utf-8",
      );
      expect(res.send).toHaveBeenCalledWith("tok/en");
      expect(calendarService.handleOutlookWebhook).not.toHaveBeenCalled();
    });

    it("processes a valid single calendar notification via handleOutlookWebhook", async () => {
      const res = makeRes();
      const body = {
        value: [calendarItem()],
      } as unknown as OutlookWebhookNotificationDto;

      await controller.handleCalendarWebhook("", body, req, res);

      expect(calendarService.handleOutlookWebhook).toHaveBeenCalledTimes(1);
      const [mapped] = calendarService.handleOutlookWebhook.mock.calls[0];
      expect(mapped).toMatchObject({
        subscriptionId: "sub-1",
        changeType: "updated",
      });
      expect(res.json).toHaveBeenCalledWith(
        expect.objectContaining({ success: true }),
      );
    });

    it("returns 202 for a batch of >2 notifications", async () => {
      const res = makeRes();
      const body = {
        value: [
          calendarItem({ subscriptionId: "s1" }),
          calendarItem({ subscriptionId: "s2" }),
          calendarItem({ subscriptionId: "s3" }),
        ],
      } as unknown as OutlookWebhookNotificationDto;

      await controller.handleCalendarWebhook("", body, req, res);

      expect(res.status).toHaveBeenCalledWith(202);
    });
  });

  describe("honors WebhookClientStateGuard verdict (stop processing)", () => {
    it("does not process an item the guard marked invalid, and still returns 200", async () => {
      const res = makeRes();
      const body = {
        value: [calendarItem()],
      } as unknown as OutlookWebhookNotificationDto;
      const guardedReq = reqWithValidation({
        valid: false,
        invalidItems: [
          { index: 0, valid: false, reason: "client_state_mismatch" },
        ],
      });

      await controller.handleCalendarWebhookNotification(
        "",
        body,
        guardedReq,
        res,
      );

      expect(calendarService.handleOutlookWebhookV2).not.toHaveBeenCalled();
      expect(res.status).not.toHaveBeenCalledWith(500);
      expect(res.json).toHaveBeenCalledWith(
        expect.objectContaining({
          success: true,
          message: expect.stringContaining("Rejected"),
        }),
      );
    });

    it("processes only the authorized items in a mixed batch", async () => {
      const res = makeRes();
      const body = {
        value: [
          calendarItem({ subscriptionId: "forged" }),
          calendarItem({ subscriptionId: "legit" }),
        ],
      } as unknown as OutlookWebhookNotificationDto;
      const guardedReq = reqWithValidation({
        valid: false,
        invalidItems: [
          { index: 0, valid: false, reason: "unknown_subscription" },
        ],
      });

      await controller.handleCalendarWebhookNotification(
        "",
        body,
        guardedReq,
        res,
      );

      expect(calendarService.handleOutlookWebhookV2).toHaveBeenCalledTimes(1);
      const [mapped] = calendarService.handleOutlookWebhookV2.mock.calls[0];
      expect(mapped).toMatchObject({ subscriptionId: "legit" });
    });

    it("legacy /webhook also skips invalid items", async () => {
      const res = makeRes();
      const body = {
        value: [calendarItem()],
      } as unknown as OutlookWebhookNotificationDto;
      const guardedReq = reqWithValidation({
        valid: false,
        invalidItems: [
          { index: 0, valid: false, reason: "missing_subscription_id" },
        ],
      });

      await controller.handleCalendarWebhook("", body, guardedReq, res);

      expect(calendarService.handleOutlookWebhook).not.toHaveBeenCalled();
      expect(res.json).toHaveBeenCalledWith(
        expect.objectContaining({ success: true }),
      );
    });
  });
});
