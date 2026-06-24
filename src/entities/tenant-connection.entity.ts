/**
 * Re-export MicrosoftTenant as TenantConnection for backward compatibility.
 *
 * Some services and controllers use the TenantConnection naming convention,
 * while the actual entity uses MicrosoftTenant. This file provides the alias.
 */
export { MicrosoftTenant as TenantConnection } from './microsoft-tenant.entity';
