/**
 * Re-export MicrosoftTenantStatus as TenantConnectionStatus for backward compatibility.
 *
 * Some services and controllers use the TenantConnectionStatus naming convention,
 * while the actual enum uses MicrosoftTenantStatus. This file provides the alias.
 */
export { MicrosoftTenantStatus as TenantConnectionStatus } from './microsoft-tenant-status.enum';
