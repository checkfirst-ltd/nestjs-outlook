import { Module, ConfigurableModuleBuilder } from "@nestjs/common";
import { TypeOrmModule } from "@nestjs/typeorm";
import { ScheduleModule } from "@nestjs/schedule";
import { EventEmitterModule } from "@nestjs/event-emitter";
import { MicrosoftAuthService } from "./services/auth/microsoft-auth.service";
import { MicrosoftAuthController } from "./controllers/microsoft-auth.controller";
import { CalendarController } from "./controllers/calendar.controller";
import { EmailController } from "./controllers/email.controller";
import { OutlookWebhookSubscription } from "./entities/outlook-webhook-subscription.entity";
import { OutlookWebhookSubscriptionRepository } from "./repositories/outlook-webhook-subscription.repository";
import { MICROSOFT_CONFIG } from "./constants";
import { MicrosoftOutlookConfig } from "./interfaces/config/outlook-config.interface";
import { MicrosoftCsrfToken } from "./entities/csrf-token.entity";
import { MicrosoftCsrfTokenRepository } from "./repositories/microsoft-csrf-token.repository";
import { CalendarService } from "./services/calendar/calendar.service";
import { EmailService } from "./services/email/email.service";
import { MicrosoftUser } from "./entities/microsoft-user.entity";
import { OutlookDeltaLink } from "./entities/delta-link.entity";
import { OutlookDeltaLinkRepository } from "./repositories/outlook-delta-link.repository";
import { DeltaSyncService } from "./services/shared/delta-sync.service";

export const { ConfigurableModuleClass, MODULE_OPTIONS_TOKEN } =
  new ConfigurableModuleBuilder<MicrosoftOutlookConfig>()
    .setClassMethodName("forRoot")
    .build();

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
    OutlookWebhookSubscriptionRepository,
    MicrosoftCsrfTokenRepository,
    CalendarService,
    EmailService,
    MicrosoftAuthService,
    OutlookDeltaLinkRepository,
    DeltaSyncService,
  ],
  exports: [
    CalendarService,
    EmailService,
    MicrosoftAuthService,
    OutlookDeltaLinkRepository,
    DeltaSyncService,
  ],
})
export class MicrosoftOutlookModule extends ConfigurableModuleClass {}
