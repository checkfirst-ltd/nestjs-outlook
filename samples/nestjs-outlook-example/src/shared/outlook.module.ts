import { Module, Global } from '@nestjs/common';
import { ConfigModule, ConfigService } from '@nestjs/config';
import {
  MicrosoftOutlookModule,
  MicrosoftOutlookConfig,
  AppOnlyAuthConfig,
  CertificateAuthConfig,
} from '@checkfirst/nestjs-outlook';
import { getRequiredConfig } from './config.utils';

/**
 * Build app-only authentication config when tenant ID is available.
 * Supports both certificate-based auth (recommended) and client secret fallback.
 */
function buildAppOnlyConfig(configService: ConfigService): AppOnlyAuthConfig | undefined {
  const tenantId = configService.get<string>('MICROSOFT_TENANT_ID');

  // App-only auth requires a tenant ID
  if (!tenantId) {
    return undefined;
  }

  // Check for certificate-based auth (more secure, recommended for production)
  const certThumbprint = configService.get<string>('MICROSOFT_CERT_THUMBPRINT');
  const certKeyPath = configService.get<string>('MICROSOFT_CERT_KEY_PATH');
  const certKeyBase64 = configService.get<string>('MICROSOFT_CERT_KEY_BASE64');

  let certificate: CertificateAuthConfig | undefined;

  if (certThumbprint && (certKeyPath || certKeyBase64)) {
    certificate = {
      thumbprint: certThumbprint,
      privateKeyPath: certKeyPath,
      privateKeyBase64: certKeyBase64,
    };
  }

  return {
    enabled: true,
    tenantId,
    certificate,
  };
}

@Global()
@Module({
  imports: [
    MicrosoftOutlookModule.forRootAsync({
      imports: [ConfigModule],
      inject: [ConfigService],
      useFactory: (configService: ConfigService): MicrosoftOutlookConfig => ({
        clientId: getRequiredConfig(configService, 'MICROSOFT_CLIENT_ID'),
        clientSecret: getRequiredConfig(configService, 'MICROSOFT_CLIENT_SECRET'),
        redirectPath: configService.get('MICROSOFT_REDIRECT_PATH', 'auth/microsoft/callback'),
        backendBaseUrl: getRequiredConfig(configService, 'BACKEND_BASE_URL'),
        basePath: getRequiredConfig(configService, 'MICROSOFT_BASE_PATH'),
        // App-only authentication for enterprise tenant access
        appOnly: buildAppOnlyConfig(configService),
      }),
    }),
  ],
  exports: [MicrosoftOutlookModule],
})
export class OutlookSharedModule {} 