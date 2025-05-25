import { Module } from '@nestjs/common';
import { AuthController } from './auth.controller';
import { MicrosoftOutlookModule } from '@checkfirst/nestjs-outlook';
import { ConfigModule, ConfigService } from '@nestjs/config';
import { getRequiredConfig } from '../shared/config.utils';

@Module({
  imports: [
    // Initialize MicrosoftOutlookModule for this feature module
    MicrosoftOutlookModule.forRootAsync({
      imports: [ConfigModule],
      inject: [ConfigService],
      useFactory: (configService: ConfigService) => ({
        clientId: getRequiredConfig(configService, 'MICROSOFT_CLIENT_ID'),
        clientSecret: getRequiredConfig(configService, 'MICROSOFT_CLIENT_SECRET'),
        redirectPath: configService.get('MICROSOFT_REDIRECT_PATH', 'auth/microsoft/callback'),
        backendBaseUrl: getRequiredConfig(configService, 'BACKEND_BASE_URL'),
        basePath: configService.get('MICROSOFT_BASE_PATH'),
      }),
    }),
  ],
  controllers: [AuthController],
})
export class AuthModule {} 