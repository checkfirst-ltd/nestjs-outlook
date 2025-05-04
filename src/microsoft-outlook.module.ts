import { Module, ConfigurableModuleBuilder } from '@nestjs/common';
import { TypeOrmModule } from '@nestjs/typeorm';
import { ScheduleModule } from '@nestjs/schedule';
import { EventEmitterModule } from '@nestjs/event-emitter';
import { OutlookService } from './services/outlook.service';
import { MicrosoftAuthService } from './services/microsoft-auth.service';
import { MicrosoftAuthController } from './controllers/microsoft-auth.controller';
import { OutlookController } from './controllers/outlook.controller';
import { OutlookWebhookSubscription } from './entities/outlook-webhook-subscription.entity';
import { OutlookWebhookSubscriptionRepository } from './repositories/outlook-webhook-subscription.repository';
import { MICROSOFT_CONFIG } from './constants';
import { MicrosoftOutlookConfig } from './interfaces/config/outlook-config.interface';
import { MicrosoftCsrfToken } from './entities/csrf-token.entity';
import { MicrosoftCsrfTokenRepository } from './repositories/microsoft-csrf-token.repository';

export const { ConfigurableModuleClass, MODULE_OPTIONS_TOKEN } =
  new ConfigurableModuleBuilder<MicrosoftOutlookConfig>().setClassMethodName('forRoot').build();

/**
 * Microsoft Outlook Module for interacting with Microsoft Graph API
 * This module should be imported using forRoot() or forRootAsync() to provide configuration
 */
@Module({
  imports: [
    ScheduleModule.forRoot(),
    TypeOrmModule.forFeature([OutlookWebhookSubscription, MicrosoftCsrfToken]),
    EventEmitterModule.forRoot(),
  ],
  controllers: [MicrosoftAuthController, OutlookController],
  providers: [
    {
      provide: MICROSOFT_CONFIG,
      useFactory: (options: MicrosoftOutlookConfig) => options,
      inject: [MODULE_OPTIONS_TOKEN],
    },
    OutlookWebhookSubscriptionRepository,
    MicrosoftCsrfTokenRepository,
    OutlookService,
    MicrosoftAuthService,
  ],
  exports: [OutlookService, MicrosoftAuthService],
})
export class MicrosoftOutlookModule extends ConfigurableModuleClass {}
