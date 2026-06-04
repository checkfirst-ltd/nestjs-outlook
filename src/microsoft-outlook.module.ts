import { Logger, Module, ConfigurableModuleBuilder } from "@nestjs/common";
import { TypeOrmModule } from "@nestjs/typeorm";
import { ScheduleModule } from "@nestjs/schedule";
import { EventEmitterModule } from "@nestjs/event-emitter";
import { MicrosoftAuthService } from "./services/auth/microsoft-auth.service";
import { MicrosoftAuthController } from "./controllers/microsoft-auth.controller";
import { CalendarController } from "./controllers/calendar.controller";
import { EmailController } from "./controllers/email.controller";
import { OutlookWebhookSubscription } from "./entities/outlook-webhook-subscription.entity";
import { OutlookWebhookSubscriptionRepository } from "./repositories/outlook-webhook-subscription.repository";
import {
  MICROSOFT_CONFIG,
  OUTLOOK_LOCK_STORE,
  OUTLOOK_RATE_LIMIT_STORE,
} from "./constants";
import { MicrosoftOutlookConfig } from "./interfaces/config/outlook-config.interface";
import { MicrosoftCsrfToken } from "./entities/csrf-token.entity";
import { MicrosoftCsrfTokenRepository } from "./repositories/microsoft-csrf-token.repository";
import { CalendarService } from "./services/calendar/calendar.service";
import { EmailService } from "./services/email/email.service";
import { MicrosoftUser } from "./entities/microsoft-user.entity";
import { OutlookDeltaLink } from "./entities/delta-link.entity";
import { OutlookDeltaLinkRepository } from "./repositories/outlook-delta-link.repository";
import { DeltaSyncService } from "./services/shared/delta-sync.service";
import { UserIdConverterService } from "./services/shared/user-id-converter.service";
import { LifecycleEventHandlerService } from "./services/calendar/lifecycle-event-handler.service";
import { RecurrenceService } from "./services/calendar/recurrence.service";
import { MicrosoftSubscriptionService } from "./services/subscription/microsoft-subscription.service";
import { GraphRateLimiterService } from "./services/shared/graph-rate-limiter.service";
import {
  InMemoryOutlookLockStore,
  OutlookLockStore,
  RedisOutlookLockStore,
} from "./services/shared/outlook-lock.store";
import {
  InMemoryOutlookRateLimitStore,
  OutlookRateLimitStore,
  RedisOutlookRateLimitStore,
} from "./services/shared/outlook-rate-limit.store";

export const { ConfigurableModuleClass, MODULE_OPTIONS_TOKEN } =
  new ConfigurableModuleBuilder<MicrosoftOutlookConfig>()
    .setClassMethodName("forRoot")
    .build();

const stateLogger = new Logger("MicrosoftOutlookStateBackend");

async function buildLockStore(
  options: MicrosoftOutlookConfig,
): Promise<OutlookLockStore> {
  const redisCfg = options.state?.redis;
  if (!redisCfg?.client) {
    stateLogger.log("OutlookLockStore backend: in-memory");
    return new InMemoryOutlookLockStore();
  }
  const prefix = redisCfg.keyPrefix ?? "outlook:";
  try {
    await redisCfg.client.ping();
    stateLogger.log(
      `OutlookLockStore backend: redis (keyPrefix="${prefix}")`,
    );
    return new RedisOutlookLockStore(redisCfg.client, prefix);
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    if (redisCfg.required) {
      stateLogger.error(
        `Redis PING failed and state.redis.required=true. Crashing module init: ${msg}`,
      );
      throw new Error(
        `MicrosoftOutlookModule: Redis state backend unreachable (required=true): ${msg}`,
      );
    }
    stateLogger.error(
      `Redis PING failed (required=false), falling back to in-memory OutlookLockStore: ${msg}. ` +
        `metric=outlook.state.backend.inmemory_fallback`,
    );
    return new InMemoryOutlookLockStore();
  }
}

async function buildRateLimitStore(
  options: MicrosoftOutlookConfig,
): Promise<OutlookRateLimitStore> {
  const redisCfg = options.state?.redis;
  if (!redisCfg?.client) {
    stateLogger.log("OutlookRateLimitStore backend: in-memory");
    return new InMemoryOutlookRateLimitStore();
  }
  const prefix = redisCfg.keyPrefix ?? "outlook:";
  try {
    await redisCfg.client.ping();
    stateLogger.log(
      `OutlookRateLimitStore backend: redis (keyPrefix="${prefix}")`,
    );
    return new RedisOutlookRateLimitStore(redisCfg.client, prefix);
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    if (redisCfg.required) {
      stateLogger.error(
        `Redis PING failed and state.redis.required=true. Crashing module init: ${msg}`,
      );
      throw new Error(
        `MicrosoftOutlookModule: Redis state backend unreachable (required=true): ${msg}`,
      );
    }
    stateLogger.error(
      `Redis PING failed (required=false), falling back to in-memory OutlookRateLimitStore: ${msg}. ` +
        `metric=outlook.state.backend.inmemory_fallback`,
    );
    return new InMemoryOutlookRateLimitStore();
  }
}

/**
 * Microsoft Outlook Module for interacting with Microsoft Graph API
 * This module should be imported using forRoot() or forRootAsync() to provide configuration
 */
@Module({
  imports: [
    ScheduleModule.forRoot(),
    TypeOrmModule.forFeature([
      OutlookWebhookSubscription,
      MicrosoftCsrfToken,
      MicrosoftUser,
      OutlookDeltaLink,
    ]),
    EventEmitterModule.forRoot(),
  ],
  controllers: [MicrosoftAuthController, CalendarController, EmailController],
  providers: [
    {
      provide: MICROSOFT_CONFIG,
      useFactory: (options: MicrosoftOutlookConfig) => options,
      inject: [MODULE_OPTIONS_TOKEN],
    },
    {
      provide: OUTLOOK_LOCK_STORE,
      useFactory: (options: MicrosoftOutlookConfig) => buildLockStore(options),
      inject: [MODULE_OPTIONS_TOKEN],
    },
    {
      provide: OUTLOOK_RATE_LIMIT_STORE,
      useFactory: (options: MicrosoftOutlookConfig) =>
        buildRateLimitStore(options),
      inject: [MODULE_OPTIONS_TOKEN],
    },
    OutlookWebhookSubscriptionRepository,
    MicrosoftCsrfTokenRepository,
    CalendarService,
    EmailService,
    MicrosoftAuthService,
    OutlookDeltaLinkRepository,
    DeltaSyncService,
    UserIdConverterService,
    LifecycleEventHandlerService,
    RecurrenceService,
    MicrosoftSubscriptionService,
    GraphRateLimiterService,
  ],
  exports: [
    CalendarService,
    RecurrenceService,
    EmailService,
    MicrosoftAuthService,
    OutlookDeltaLinkRepository,
    DeltaSyncService,
    UserIdConverterService,
    MicrosoftSubscriptionService,
    GraphRateLimiterService,
    OUTLOOK_LOCK_STORE,
    OUTLOOK_RATE_LIMIT_STORE,
  ],
})
export class MicrosoftOutlookModule extends ConfigurableModuleClass {}
