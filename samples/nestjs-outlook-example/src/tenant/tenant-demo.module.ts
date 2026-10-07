import { Module } from '@nestjs/common';
import { ConfigModule, ConfigService } from '@nestjs/config';
import { MicrosoftOutlookModule } from '@checkfirst/nestjs-outlook';
import { TenantDemoController } from './tenant-demo.controller';
import { getRequiredConfig } from '../shared/config.utils';

/**
 * Module demonstrating tenant-wide Microsoft 365 operations.
 *
 * This module showcases:
 * - TenantCalendarService: Access any user's calendar with app-only auth
 * - TenantUserService: Look up and map users within the tenant
 *
 * Required environment variables:
 * - MICROSOFT_CLIENT_ID: Azure AD application ID
 * - MICROSOFT_CLIENT_SECRET: Application secret (for delegated auth fallback)
 * - BACKEND_BASE_URL: Your backend URL
 * - DEMO_TENANT_ID: The Microsoft tenant ID for demo purposes
 *
 * For app-only authentication, you must also configure:
 * - Certificate files (PEM format)
 * - Certificate thumbprint
 * - Admin consent in Azure AD
 */
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
        basePath: configService.get('MICROSOFT_BASE_PATH'),
      }),
    }),
  ],
  controllers: [TenantDemoController],
})
export class TenantDemoModule {}
