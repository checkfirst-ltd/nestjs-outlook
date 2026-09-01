import { Module } from '@nestjs/common';
import { ConfigModule, ConfigService } from '@nestjs/config';
import { TypeOrmModule } from '@nestjs/typeorm';
import {
  MicrosoftOutlookModule,
  MicrosoftTenant,
  MicrosoftUser,
  AppOnlyAuthConfig,
} from '@checkfirst/nestjs-outlook';
import { TenantController } from './tenant.controller';
import { TenantService } from './tenant.service';
import { CertificateService } from './certificate.service';
import { getRequiredConfig } from '../shared/config.utils';

/**
 * Build the app-only (client-credentials) config from the shared certificate
 * environment variables. Required so that string-based token acquisition
 * (TenantUserService.listUsers, lookupUserByEmail, etc.) can resolve credentials.
 *
 * `tenantId` is only a placeholder to enable the feature — the multi-tenant
 * services always pass an explicit tenant ID, which overrides it per request.
 */
function buildAppOnlyConfig(configService: ConfigService): AppOnlyAuthConfig | undefined {
  const thumbprint = configService.get<string>('MICROSOFT_CERTIFICATE_THUMBPRINT');
  const privateKeyPath = configService.get<string>('MICROSOFT_CERTIFICATE_KEY_PATH');
  const certificatePath = configService.get<string>('MICROSOFT_CERTIFICATE_PATH');

  if (!thumbprint || !privateKeyPath) {
    return undefined;
  }

  return {
    enabled: true,
    tenantId: configService.get<string>('MICROSOFT_TENANT_ID') ?? 'common',
    certificate: { thumbprint, privateKeyPath, certificatePath },
  };
}

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
    TypeOrmModule.forFeature([MicrosoftTenant, MicrosoftUser]),
    MicrosoftOutlookModule.forRootAsync({
      imports: [ConfigModule],
      inject: [ConfigService],
      useFactory: (configService: ConfigService) => ({
        clientId: getRequiredConfig(configService, 'MICROSOFT_CLIENT_ID'),
        clientSecret: getRequiredConfig(configService, 'MICROSOFT_CLIENT_SECRET'),
        redirectPath: configService.get('MICROSOFT_REDIRECT_PATH', 'auth/microsoft/callback'),
        backendBaseUrl: getRequiredConfig(configService, 'BACKEND_BASE_URL'),
        basePath: configService.get('MICROSOFT_BASE_PATH'),
        // App-only (client-credentials) config built from the shared certificate.
        appOnly: buildAppOnlyConfig(configService),
      }),
    }),
  ],
  controllers: [TenantController],
  providers: [TenantService, CertificateService],
  exports: [TenantService],
})
export class TenantModule {}
