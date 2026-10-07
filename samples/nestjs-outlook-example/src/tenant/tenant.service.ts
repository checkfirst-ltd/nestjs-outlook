import { BadRequestException, Injectable, Logger, NotFoundException } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { InjectRepository } from '@nestjs/typeorm';
import { Repository } from 'typeorm';
import {
  TenantCalendarService,
  TenantUserService,
  AppOnlyAuthService,
  MicrosoftTenant,
  MicrosoftUser,
  MicrosoftTenantStatus,
} from '@checkfirst/nestjs-outlook';
import { CreateTenantEventDto } from './dto/create-tenant-event.dto';
import { RegisterUserMappingDto } from './dto/lookup-user.dto';
import { RegisterTenantDto } from './dto/register-tenant.dto';
import { GenerateCertificateDto } from './dto/generate-certificate.dto';
import { CertificateService } from './certificate.service';

/**
 * Service for managing tenant-wide calendar operations.
 *
 * Demonstrates how to:
 * - Look up Microsoft users by email
 * - Register user mappings (external user ID -> Microsoft user ID)
 * - Create calendar events for any user in the tenant
 * - Get calendar info for tenant users
 */
@Injectable()
export class TenantService {
  private readonly logger = new Logger(TenantService.name);

  constructor(
    private readonly tenantCalendarService: TenantCalendarService,
    private readonly tenantUserService: TenantUserService,
    private readonly appOnlyAuthService: AppOnlyAuthService,
    private readonly certificateService: CertificateService,
    private readonly configService: ConfigService,
    @InjectRepository(MicrosoftTenant)
    private readonly tenantRepository: Repository<MicrosoftTenant>,
    @InjectRepository(MicrosoftUser)
    private readonly tenantUserRepository: Repository<MicrosoftUser>,
  ) {}

  /**
   * Generate a self-signed certificate for app-only authentication.
   * `shared: true` generates the one-time Checkfirst shared-app certificate;
   * otherwise a per-tenant certificate is generated for `tenantId`.
   */
  generateCertificate(dto: GenerateCertificateDto) {
    if (dto.shared) {
      return this.certificateService.generateSharedCertificate();
    }

    if (!dto.tenantId) {
      throw new BadRequestException('tenantId is required to generate a dedicated certificate.');
    }

    return this.certificateService.generateTenantCertificate(dto.tenantId);
  }

  /**
   * Register (or update) a Microsoft 365 tenant for app-only access.
   *
   * This inserts the `microsoft_tenants` row that the admin-consent callback
   * later activates. The tenant starts in PENDING_CONSENT status; an Azure AD
   * administrator must then visit the returned `adminConsentUrl` to grant
   * tenant-wide application permissions.
   *
   * Two onboarding models are supported:
   * - `shared`: credentials default to the Checkfirst app's configured client ID
   *   and shared certificate (the customer only supplies `tenantId`).
   * - `dedicated`: the customer supplies their own `clientId` plus the per-tenant
   *   certificate fields (typically from POST /tenant/certificate/generate).
   */
  async registerTenant(dto: RegisterTenantDto) {
    const mode = dto.mode ?? 'shared';
    this.logger.log(`Registering tenant ${dto.tenantId} (mode ${mode})`);

    const credentials = this.resolveRegistrationCredentials(dto, mode);

    const existing = await this.tenantRepository.findOne({
      where: { tenantId: dto.tenantId },
    });

    const tenant = existing ?? this.tenantRepository.create({ tenantId: dto.tenantId });

    tenant.clientId = credentials.clientId;
    tenant.certificateThumbprint = credentials.certificateThumbprint;
    tenant.certificatePath = credentials.certificatePath;
    tenant.certificateKeyPath = credentials.certificateKeyPath;
    tenant.isActive = true;
    // Preserve ACTIVE status on re-registration; otherwise (re)start the consent flow.
    if (tenant.status !== MicrosoftTenantStatus.ACTIVE) {
      tenant.status = MicrosoftTenantStatus.PENDING_CONSENT;
    }

    const saved = await this.tenantRepository.save(tenant);

    // The consent callback looks the tenant up by the `state` parameter, which
    // this implementation maps to `tenantId` — so state and tenantId are both
    // the directory GUID.
    const adminConsentUrl = this.appOnlyAuthService.getAdminConsentUrl(
      saved.tenantId,
      saved.tenantId,
      saved.clientId,
    );

    return {
      id: saved.id,
      tenantId: saved.tenantId,
      clientId: saved.clientId,
      status: saved.status,
      isActive: saved.isActive,
      adminConsentUrl,
    };
  }

  /**
   * Resolve the client ID + certificate fields to persist, based on the
   * onboarding mode. In shared mode they come from the Checkfirst app config;
   * in dedicated mode they come from the request (the tenant's own app + cert).
   */
  private resolveRegistrationCredentials(
    dto: RegisterTenantDto,
    mode: 'shared' | 'dedicated',
  ): {
    clientId: string;
    certificateThumbprint: string;
    certificatePath: string | null;
    certificateKeyPath: string | null;
  } {
    if (mode === 'dedicated') {
      if (!dto.clientId) {
        throw new BadRequestException('clientId is required in "dedicated" mode.');
      }
      if (!dto.certificateThumbprint || !dto.certificateKeyPath) {
        throw new BadRequestException(
          'certificateThumbprint and certificateKeyPath are required in "dedicated" mode. ' +
            'Generate them first with POST /tenant/certificate/generate.',
        );
      }
      return {
        clientId: dto.clientId,
        certificateThumbprint: dto.certificateThumbprint,
        certificatePath: dto.certificatePath ?? null,
        certificateKeyPath: dto.certificateKeyPath,
      };
    }

    // shared mode — default everything from the Checkfirst app config.
    const clientId = this.appOnlyAuthService.getClientId();
    const certificateThumbprint = this.configService.get<string>('MICROSOFT_CERTIFICATE_THUMBPRINT');
    const certificateKeyPath = this.configService.get<string>('MICROSOFT_CERTIFICATE_KEY_PATH');
    const certificatePath = this.configService.get<string>('MICROSOFT_CERTIFICATE_PATH');

    if (!certificateThumbprint || !certificateKeyPath) {
      throw new BadRequestException(
        'Shared certificate is not configured. Run the one-time setup ' +
          '(POST /tenant/certificate/generate { "shared": true }), upload certs/shared.crt to ' +
          'the Checkfirst Azure app, then set MICROSOFT_CERTIFICATE_THUMBPRINT, ' +
          'MICROSOFT_CERTIFICATE_KEY_PATH and MICROSOFT_CERTIFICATE_PATH and restart.',
      );
    }

    return {
      clientId,
      certificateThumbprint,
      certificatePath: certificatePath ?? null,
      certificateKeyPath,
    };
  }

  /**
   * Get the active tenant configuration.
   * In production, you would likely have multiple tenants and select based on context.
   */
  async getActiveTenant(): Promise<MicrosoftTenant> {
    const tenant = await this.tenantRepository.findOne({
      where: {
        isActive: true,
        status: MicrosoftTenantStatus.ACTIVE,
      },
    });

    if (!tenant) {
      throw new NotFoundException(
        'No active tenant configured. Please complete the admin consent flow first.',
      );
    }

    return tenant;
  }

  /**
   * List users in the active tenant (app-only). Optional OData `$filter`.
   */
  async listUsers(filter?: string) {
    const tenant = await this.getActiveTenant();
    this.logger.log(`Listing users in tenant ${tenant.tenantId}${filter ? ` (filter: ${filter})` : ''}`);
    return this.tenantUserService.listUsers(tenant.tenantId, filter ? { filter } : undefined);
  }

  /**
   * Get a single Microsoft user in the active tenant by UPN/email or object ID.
   */
  async getUser(userIdOrUpn: string) {
    const tenant = await this.getActiveTenant();
    this.logger.log(`Getting user ${userIdOrUpn} in tenant ${tenant.tenantId}`);

    // A value containing "@" is treated as a UPN/email; otherwise as an object ID.
    const user = userIdOrUpn.includes('@')
      ? await this.tenantUserService.lookupUserByUpn(tenant.tenantId, userIdOrUpn)
      : await this.tenantUserService.getUserById(tenant.tenantId, userIdOrUpn);

    if (!user) {
      throw new NotFoundException(`User not found: ${userIdOrUpn}`);
    }

    return user;
  }

  /**
   * List the calendars for a user in the active tenant (app-only).
   */
  async getUserCalendars(userId: string) {
    const data = await this.graphGet(
      `/users/${encodeURIComponent(userId)}/calendars?$select=id,name,owner,canEdit,isDefaultCalendar`,
    );
    return { calendars: data.value ?? [] };
  }

  /**
   * List upcoming/recent events for a user in the active tenant (app-only).
   */
  async getUserEvents(userId: string) {
    const data = await this.graphGet(
      `/users/${encodeURIComponent(userId)}/events?$select=id,subject,start,end,location,organizer&$top=25`,
    );
    return { events: data.value ?? [] };
  }

  /**
   * Perform a GET against Microsoft Graph using an app-only token for the
   * active tenant. Credentials are resolved from the tenant entity, so this
   * works for both shared and dedicated certificate models.
   */
  private async graphGet(path: string): Promise<{ value?: unknown[] }> {
    const tenant = await this.getActiveTenant();
    const token = await this.appOnlyAuthService.getAccessToken(tenant);

    const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    });

    const data = (await response.json()) as { value?: unknown[]; error?: { message?: string } };
    if (!response.ok) {
      throw new BadRequestException(
        data.error?.message ?? `Microsoft Graph request failed (${response.status})`,
      );
    }
    return data;
  }

  /**
   * Look up a Microsoft user by email address within the tenant.
   */
  async lookupUser(email: string) {
    const tenant = await this.getActiveTenant();

    this.logger.log(`Looking up user by email: ${email} in tenant ${tenant.tenantId}`);

    const user = await this.tenantUserService.lookupUserByEmail(
      tenant.tenantId,
      email,
    );

    if (!user) {
      throw new NotFoundException(`User not found with email: ${email}`);
    }

    return {
      microsoftUserId: user.microsoftUserId,
      userPrincipalName: user.userPrincipalName,
      displayName: user.displayName,
      email: user.email,
    };
  }

  /**
   * Register a mapping between an external user ID and a Microsoft user.
   */
  async registerUserMapping(dto: RegisterUserMappingDto) {
    const tenant = await this.getActiveTenant();

    this.logger.log(
      `Registering user mapping: ${dto.externalUserId} -> ${dto.email} in tenant ${tenant.tenantId}`,
    );

    // Look up the Microsoft user to get their ID
    const msUser = await this.tenantUserService.lookupUserByEmail(
      tenant.tenantId,
      dto.email,
    );

    if (!msUser) {
      throw new NotFoundException(`Microsoft user not found with email: ${dto.email}`);
    }

    // Check if mapping already exists
    let tenantUser = await this.tenantUserRepository.findOne({
      where: {
        tenant: { id: tenant.id },
        externalUserId: dto.externalUserId,
      },
    });

    if (tenantUser) {
      // Update existing mapping
      tenantUser.microsoftUserId = msUser.microsoftUserId;
      tenantUser.userPrincipalName = msUser.userPrincipalName;
    } else {
      // Create new mapping
      tenantUser = this.tenantUserRepository.create({
        tenant,
        externalUserId: dto.externalUserId,
        microsoftUserId: msUser.microsoftUserId,
        userPrincipalName: msUser.userPrincipalName,
        isActive: true,
      });
    }

    await this.tenantUserRepository.save(tenantUser);

    return {
      externalUserId: dto.externalUserId,
      microsoftUserId: msUser.microsoftUserId,
      userPrincipalName: msUser.userPrincipalName,
      displayName: msUser.displayName,
    };
  }

  /**
   * Get the Microsoft user ID for an external user.
   */
  private async getMicrosoftUserId(externalUserId: string): Promise<{ tenantId: string; microsoftUserId: string }> {
    const tenant = await this.getActiveTenant();

    const tenantUser = await this.tenantUserRepository.findOne({
      where: {
        tenant: { id: tenant.id },
        externalUserId,
        isActive: true,
      },
    });

    if (!tenantUser || !tenantUser.microsoftUserId) {
      throw new NotFoundException(
        `No user mapping found for external user ID: ${externalUserId}. ` +
        'Register the user mapping first using POST /tenant/users/register.',
      );
    }

    return {
      tenantId: tenant.tenantId,
      microsoftUserId: tenantUser.microsoftUserId,
    };
  }

  /**
   * Get the default calendar ID for a tenant user.
   */
  async getDefaultCalendarId(externalUserId: string) {
    const { tenantId, microsoftUserId } = await this.getMicrosoftUserId(externalUserId);

    this.logger.log(`Getting default calendar for user ${microsoftUserId}`);

    const calendarId = await this.tenantCalendarService.getDefaultCalendarId(
      tenantId,
      microsoftUserId,
    );

    return {
      externalUserId,
      microsoftUserId,
      calendarId,
    };
  }

  /**
   * Get a specific event by ID for a tenant user.
   */
  async getEvent(externalUserId: string, eventId: string) {
    const { tenantId, microsoftUserId } = await this.getMicrosoftUserId(externalUserId);

    this.logger.log(`Getting event ${eventId} for user ${microsoftUserId}`);

    const event = await this.tenantCalendarService.getEventById(
      tenantId,
      microsoftUserId,
      eventId,
    );

    return event;
  }

  /**
   * Create a calendar event for a tenant user.
   */
  async createEvent(dto: CreateTenantEventDto) {
    const { tenantId, microsoftUserId } = await this.getMicrosoftUserId(dto.externalUserId);

    this.logger.log(`Creating event for user ${microsoftUserId}: ${dto.subject}`);

    // Get the default calendar ID
    const calendarId = await this.tenantCalendarService.getDefaultCalendarId(
      tenantId,
      microsoftUserId,
    );

    // Build the event object matching the Microsoft Graph Event type
    const eventData = {
      subject: dto.subject,
      start: {
        dateTime: dto.startDateTime,
        timeZone: 'UTC',
      },
      end: {
        dateTime: dto.endDateTime,
        timeZone: 'UTC',
      },
      body: dto.body ? {
        contentType: 'html' as const,
        content: dto.body,
      } : undefined,
      location: dto.location ? {
        displayName: dto.location,
      } : undefined,
    };

    const result = await this.tenantCalendarService.createEvent(
      eventData,
      tenantId,
      microsoftUserId,
      calendarId,
    );

    return result;
  }

  /**
   * Get tenant status summary.
   */
  async getTenantStatus() {
    const tenants = await this.tenantRepository.find();
    const userCount = await this.tenantUserRepository.count({
      where: { isActive: true },
    });

    return {
      tenants: tenants.map(t => ({
        tenantId: t.tenantId,
        status: t.status,
        isActive: t.isActive,
        adminConsentGrantedAt: t.adminConsentGrantedAt,
      })),
      totalMappedUsers: userCount,
    };
  }
}
