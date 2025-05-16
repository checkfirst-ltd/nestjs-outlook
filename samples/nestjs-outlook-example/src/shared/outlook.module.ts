import { Module, Global } from '@nestjs/common';
import { ConfigModule, ConfigService } from '@nestjs/config';
import { MicrosoftOutlookModule } from '@checkfirst/nestjs-outlook';

@Global()
@Module({
  imports: [
    MicrosoftOutlookModule.forRootAsync({
      imports: [ConfigModule],
      inject: [ConfigService],
      useFactory: (configService: ConfigService) => ({
        clientId: configService.get('MICROSOFT_CLIENT_ID'),
        clientSecret: configService.get('MICROSOFT_CLIENT_SECRET'),
        redirectPath: configService.get('MICROSOFT_REDIRECT_PATH', 'auth/microsoft/callback'),
        backendBaseUrl: configService.get('BACKEND_BASE_URL'),
        basePath: configService.get('MICROSOFT_BASE_PATH'),
      }),
    }),
  ],
  exports: [MicrosoftOutlookModule],
})
export class OutlookSharedModule {} 