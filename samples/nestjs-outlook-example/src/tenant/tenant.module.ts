import { Module } from '@nestjs/common';
import { ConfigModule, ConfigService } from '@nestjs/config';
import { TypeOrmModule } from '@nestjs/typeorm';
import {
  MicrosoftOutlookModule,
  MicrosoftTenant,
  MicrosoftTenantUser,
} from '@checkfirst/nestjs-outlook';
import { TenantController } from './tenant.controller';
import { TenantService } from './tenant.service';
import { getRequiredConfig } from '../shared/config.utils';

/**
 * Module for enterprise tenant calendar operations using app-only authentication.
 *
 * This module demonstrates how to use TenantCalendarService and TenantUserService
 * to access calendars across all users in a Microsoft 365 tenant without requiring
 * individual user consent.
 *
 * Prerequisites:
 * - Azure AD app registration with Application permissions
 * - Admin consent granted for the tenant
 * - Certificate-based authentication configured
 */
@Module({
  imports: [
    TypeOrmModule.forFeature([MicrosoftTenant, MicrosoftTenantUser]),
    MicrosoftOutlookModule.forRootAsync({
      imports: [ConfigModule],
      inject: [ConfigService],
      useFactory: (configService: ConfigService) => ({
        clientId: getRequiredConfig(configService, 'MICROSOFT_CLIENT_ID'),
        clientSecret: getRequiredConfig(configService, 'MICROSOFT_CLIENT_SECRET'),
        redirectPath: configService.get('MICROSOFT_REDIRECT_PATH', 'auth/microsoft/callback'),
        backendBaseUrl: getRequiredConfig(configService, 'BACKEND_BASE_URL'),
        basePath: configService.get('MICROSOFT_BASE_PATH'),
        // Tenant-specific configuration for app-only auth
        tenant: {
          tenantId: configService.get('MICROSOFT_TENANT_ID'),
          certificatePath: configService.get('MICROSOFT_CERTIFICATE_PATH'),
          certificateKeyPath: configService.get('MICROSOFT_CERTIFICATE_KEY_PATH'),
          certificateThumbprint: configService.get('MICROSOFT_CERTIFICATE_THUMBPRINT'),
        },
      }),
    }),
  ],
  controllers: [TenantController],
  providers: [TenantService],
  exports: [TenantService],
})
export class TenantModule {}
