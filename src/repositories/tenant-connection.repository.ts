/**
 * Re-export MicrosoftTenantRepository as TenantConnectionRepository for backward compatibility.
 *
 * Some services and controllers use the TenantConnectionRepository naming convention,
 * while the actual repository uses MicrosoftTenantRepository. This file provides the alias.
 */
export { MicrosoftTenantRepository as TenantConnectionRepository } from './microsoft-tenant.repository';
