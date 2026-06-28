import { Injectable, Logger } from '@nestjs/common';
import { createHash } from 'crypto';
import * as fs from 'fs';
import { join } from 'path';
import * as forge from 'node-forge';

/**
 * Result of generating a self-signed certificate + keypair.
 */
export interface GeneratedCertificate {
  /** x5t#S256 thumbprint (base64url SHA-256 of the certificate DER). */
  thumbprint: string;
  /** PEM-encoded public certificate (the `.cer`/`.crt` to upload to Azure). */
  certificatePem: string;
  /** Absolute path to the persisted private key PEM. */
  keyPath: string;
  /** Absolute path to the persisted certificate PEM. */
  certPath: string;
}

/**
 * Generates self-signed X.509 certificates for app-only (client-credentials)
 * Microsoft Graph authentication.
 *
 * DEMO ONLY: private keys are written to disk as unencrypted PEM under `certs/`.
 * A production system should generate/store keys in a KMS/HSM (e.g. Azure Key
 * Vault) and never persist unencrypted key material.
 */
@Injectable()
export class CertificateService {
  private readonly logger = new Logger(CertificateService.name);
  private readonly certsDir = join(process.cwd(), 'certs');
  /** Certificate validity in years. */
  private readonly validityYears = 2;

  /**
   * Generate a dedicated certificate for a specific tenant. The private key and
   * certificate are persisted as `certs/<tenantId>.{key,crt}`.
   */
  generateTenantCertificate(tenantId: string): GeneratedCertificate {
    this.logger.log(`Generating dedicated certificate for tenant ${tenantId}`);
    return this.generateAndPersist(`checkfirst-tenant-${tenantId}`, tenantId);
  }

  /**
   * Generate the one-time shared certificate for the Checkfirst-owned multi-tenant
   * app registration. Persisted as `certs/shared.{key,crt}`. Operator-only.
   */
  generateSharedCertificate(): GeneratedCertificate & { setupHint: string } {
    this.logger.log('Generating shared Checkfirst app certificate');
    const result = this.generateAndPersist('checkfirst-shared-app', 'shared');
    return {
      ...result,
      setupHint:
        'Upload the certificate (certs/shared.crt) to the Checkfirst-owned Azure app ' +
        'registration once, then set MICROSOFT_CERTIFICATE_PATH, MICROSOFT_CERTIFICATE_KEY_PATH ' +
        'and MICROSOFT_CERTIFICATE_THUMBPRINT to the returned values and restart the app.',
    };
  }

  /**
   * Build a self-signed certificate + RSA keypair and compute its x5t#S256 thumbprint.
   */
  private buildSelfSigned(commonName: string): {
    certPem: string;
    keyPem: string;
    thumbprint: string;
  } {
    const keys = forge.pki.rsa.generateKeyPair({ bits: 2048 });
    const cert = forge.pki.createCertificate();

    cert.publicKey = keys.publicKey;
    cert.serialNumber = this.randomSerial();

    const now = new Date();
    cert.validity.notBefore = now;
    cert.validity.notAfter = new Date(
      now.getFullYear() + this.validityYears,
      now.getMonth(),
      now.getDate(),
    );

    const attrs = [{ name: 'commonName', value: commonName }];
    cert.setSubject(attrs);
    cert.setIssuer(attrs);
    cert.setExtensions([{ name: 'basicConstraints', cA: false }]);

    // Self-sign with SHA-256.
    cert.sign(keys.privateKey, forge.md.sha256.create());

    const certPem = forge.pki.certificateToPem(cert);
    const keyPem = forge.pki.privateKeyToPem(keys.privateKey);

    // x5t#S256 = base64url(SHA-256(DER(cert))). Byte-identical to the openssl
    // pipeline `x509 -outform DER | dgst -sha256 -binary | base64 | tr +/ -_ | tr -d =`.
    const der = forge.asn1.toDer(forge.pki.certificateToAsn1(cert)).getBytes();
    const thumbprint = createHash('sha256')
      .update(Buffer.from(der, 'binary'))
      .digest('base64url');

    return { certPem, keyPem, thumbprint };
  }

  /**
   * Build a self-signed cert and persist key + cert to the certs directory.
   */
  private generateAndPersist(commonName: string, fileBase: string): GeneratedCertificate {
    const { certPem, keyPem, thumbprint } = this.buildSelfSigned(commonName);

    fs.mkdirSync(this.certsDir, { recursive: true });
    const keyPath = join(this.certsDir, `${fileBase}.key`);
    const certPath = join(this.certsDir, `${fileBase}.crt`);

    // Restrict the private key to the owner (demo-grade protection).
    fs.writeFileSync(keyPath, keyPem, { mode: 0o600 });
    fs.writeFileSync(certPath, certPem);

    this.logger.log(`Certificate written to ${certPath} (thumbprint ${thumbprint})`);

    return { thumbprint, certificatePem: certPem, keyPath, certPath };
  }

  /**
   * Generate a positive 1-byte-aligned hex serial number for the certificate.
   */
  private randomSerial(): string {
    const bytes = forge.random.getBytesSync(16);
    // Ensure the high bit is clear so the integer is interpreted as positive.
    return '00' + forge.util.bytesToHex(bytes);
  }
}
