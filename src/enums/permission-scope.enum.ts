/**
 * Generic permission scopes that can be requested by the host application
 * These are provider-agnostic and will be mapped to specific provider scopes
 */
export enum PermissionScope {
  CALENDAR_READ = 'CALENDAR_READ',
  CALENDAR_WRITE = 'CALENDAR_WRITE',
  EMAIL_READ = 'EMAIL_READ',
  EMAIL_WRITE = 'EMAIL_WRITE',
  EMAIL_SEND = 'EMAIL_SEND',
} 