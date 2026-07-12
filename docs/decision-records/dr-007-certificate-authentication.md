---
dep:
  type: decision-record
  audience: [library-contributor, ai-agent]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/interfaces/config/outlook-config.interface.ts
    - src/services/auth/app-only-auth.service.ts
  tags: [decision, auth, certificate, security, app-only]
  links:
    - target: ../reference/app-only-auth-service.md
      rel: DECIDES
    - target: ../reference/configuration.md
      rel: DECIDES
    - target: dr-006-dual-auth-architecture.md
      rel: EXTENDS
---

# DR-007: Certificate Authentication for App-Only Mode

## Context

App-only authentication requires proving the application's identity to Azure AD. Microsoft
supports two credential types:

1. **Client secrets:** Simple strings that are easy to configure but must be rotated
   manually and can be leaked if stored insecurely.

2. **Certificates:** X.509 certificates where the private key signs a JWT assertion,
   proving possession without transmitting the secret.

Enterprise customers with security requirements often mandate certificate-based
authentication. The module needed to support both options while guiding users toward
the more secure choice.

## Decision

Support both client secrets and certificates for app-only authentication, with certificates
taking precedence when configured:

```typescript
appOnly: {
  enabled: true,
  tenantId: 'tenant-id',

  // Option 1: Client secret (simpler)
  // Uses clientSecret from parent config

  // Option 2: Certificate (more secure, takes precedence)
  certificate: {
    thumbprint: 'CERT_THUMBPRINT_HEX',
    privateKey: '-----BEGIN PRIVATE KEY-----\n...',
  },
}
```

When `certificate` is configured:
1. Generate a JWT assertion with claims required by Azure AD
2. Sign the assertion using the private key
3. Include the certificate thumbprint in the JWT header
4. Exchange the assertion for an access token via the OAuth2 token endpoint

The module never stores or logs the private key. It accepts the key as a string (PEM format)
loaded by the host from secure storage (environment variable, secrets manager, HSM).

## JWT assertion structure

The signed assertion follows Microsoft's client assertion format:

**Header:**
```json
{
  "alg": "RS256",
  "typ": "JWT",
  "x5t": "<base64url-encoded certificate thumbprint>"
}
```

**Payload:**
```json
{
  "aud": "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
  "iss": "<client_id>",
  "sub": "<client_id>",
  "jti": "<unique identifier>",
  "nbf": <current timestamp>,
  "exp": <current timestamp + 10 minutes>
}
```

## Alternatives considered

### Certificate file path

Accept a file path to the certificate/key rather than the PEM string directly.

**Rejected because:**
- File paths leak implementation details into configuration
- Harder to use with secrets managers that provide values, not files
- Container deployments often prefer environment variables or mounted secrets

### PKCS#12/PFX support

Accept `.pfx` bundles that contain both certificate and private key.

**Rejected because:**
- Adds dependency for PFX parsing
- Environment variable transport is awkward for binary formats
- PEM is the standard for programmatic key handling

### Hardware Security Module (HSM) integration

Direct integration with Azure Key Vault or HSMs for key storage.

**Rejected because:**
- Adds significant complexity and dependencies
- Host applications often have their own key management infrastructure
- Can be implemented by the host loading the key from HSM before passing to module

## Consequences

### Positive

- **Security improvement:** Certificates provide asymmetric authentication — the secret
  (private key) never leaves the host; only the signature is transmitted
- **Compliance friendly:** Meets enterprise security requirements that forbid shared secrets
- **Rotation patterns:** Certificates have built-in expiry; new certs can be deployed
  alongside old ones during rotation
- **Clear precedence:** No ambiguity about which credential is used when both are present

### Negative

- **Configuration complexity:** Certificates require more setup than client secrets
- **Key management burden:** Host must securely store and rotate private keys
- **Debugging difficulty:** Certificate errors (wrong thumbprint, key mismatch) are harder
  to diagnose than secret errors

### Operational

- Private key must be available at module initialization
- Certificate thumbprint must match the certificate registered in Azure AD
- Assertion JWTs are short-lived (10 minutes) and generated per token request
- Token caching reduces frequency of assertion generation

## Security considerations

1. **Never log the private key.** The module accepts it as a parameter but does not
   include it in any log output or error messages.

2. **Validate certificate expiry.** The module does not validate the certificate's
   expiry date — this is the host's responsibility during deployment.

3. **Thumbprint verification.** Azure AD verifies the thumbprint matches a registered
   certificate. Mismatches fail fast with a clear error.

4. **Key rotation.** Azure AD supports multiple certificates per app registration.
   Deploy the new certificate, update the module configuration, then remove the old
   certificate from Azure AD.

## Review trigger

Revisit if:
- Microsoft changes the client assertion format or signing requirements
- Demand emerges for PKCS#12 or HSM integration
- A security audit identifies weaknesses in the current implementation
