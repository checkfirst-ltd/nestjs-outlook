import { Logger } from '@nestjs/common';
import { MicrosoftSubscriptionService } from './microsoft-subscription.service';
import { OutlookEventTypes } from '../../enums/event-types.enum';
import { MicrosoftOutlookConfig } from '../../interfaces/config/outlook-config.interface';

/**
 * Focused tests for the catch-up reconcile primitive and staleness config.
 *
 * triggerCatchUpReconcile is the single safe recovery entry point used by the missed-lifecycle
 * handler, the health-check stale detector, and the renew/recreate paths. It emits a synthetic
 * EVENT_NOTIFICATION that drives the cursor-gated reconcile in calendar-hub (NOT the legacy
 * at-most-once syncDeltaChanges).
 *
 * Note: Subscription renewal is now handled proactively by SubscriptionRenewalWorker,
 * not by the health check. This service only handles staleness detection.
 */
describe('MicrosoftSubscriptionService — catch-up + staleness config', () => {
  let emit: jest.Mock;

  const make = (subscription?: MicrosoftOutlookConfig['subscription']) => {
    emit = jest.fn();
    const config = {
      clientId: '',
      clientSecret: '',
      redirectPath: '',
      backendBaseUrl: '',
      subscription,
    } as MicrosoftOutlookConfig;
    return new MicrosoftSubscriptionService(
      {} as any, // microsoftAuthService
      {} as any, // webhookSubscriptionRepository
      { emit } as any, // eventEmitter
      config,
      {} as any, // microsoftUserRepository
      {} as any, // userIdConverter
    );
  };

  beforeEach(() => {
    jest.spyOn(Logger.prototype, 'log').mockImplementation(() => undefined);
    jest.spyOn(Logger.prototype, 'warn').mockImplementation(() => undefined);
  });
  afterEach(() => jest.restoreAllMocks());

  it('triggerCatchUpReconcile emits an EVENT_NOTIFICATION that drives a full per-user reconcile', () => {
    const svc = make();

    svc.triggerCatchUpReconcile(1789, 'lifecycle-missed');

    expect(emit).toHaveBeenCalledTimes(1);
    expect(emit).toHaveBeenCalledWith(OutlookEventTypes.EVENT_NOTIFICATION, {
      userId: 1789,
      resource: { id: 'catchup:lifecycle-missed' },
      changeType: 'updated', // non-'deleted' -> consumer runs the reconcile branch
    });
  });

  it('triggerCatchUpReconcile never throws into the caller if the event bus fails', () => {
    const svc = make();
    emit.mockImplementationOnce(() => {
      throw new Error('bus down');
    });

    expect(() => {
      svc.triggerCatchUpReconcile(1, 'renew');
    }).not.toThrow();
  });

  it('defaults staleNotificationThresholdHours to 24h', () => {
    const svc = make();

    expect((svc as any).staleNotificationThresholdHours).toBe(24);
  });

  it('honors staleNotificationThresholdHours config override', () => {
    const svc = make({ staleNotificationThresholdHours: 6 });

    expect((svc as any).staleNotificationThresholdHours).toBe(6);
  });
});
