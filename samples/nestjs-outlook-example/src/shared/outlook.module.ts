import { Module, Global } from '@nestjs/common';
import { ConfigModule, ConfigService } from '@nestjs/config';
import { MicrosoftOutlookModule } from '@checkfirst/nestjs-outlook';
import { getRequiredConfig } from './config.utils';

@Global()
@Module({
  imports: [
    MicrosoftOutlookModule.forRootAsync({
      imports: [ConfigModule],
      inject: [ConfigService],
      useFactory: (configService: ConfigService) => ({
        clientId: getRequiredConfig(configService, 'MICROSOFT_CLIENT_ID'),
        clientSecret: getRequiredConfig(configService, 'MICROSOFT_CLIENT_SECRET'),
        redirectPath: configService.get('MICROSOFT_REDIRECT_PATH', 'auth/microsoft/callback'),
        backendBaseUrl: getRequiredConfig(configService, 'BACKEND_BASE_URL'),
        basePath: getRequiredConfig(configService, 'MICROSOFT_BASE_PATH'),
      }),
    }),
  ],
  exports: [MicrosoftOutlookModule],
})
export class OutlookSharedModule {} 