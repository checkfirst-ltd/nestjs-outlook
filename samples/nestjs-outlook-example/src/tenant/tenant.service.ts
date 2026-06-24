import { Injectable, Logger, NotFoundException } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { Repository } from 'typeorm';
import {
  TenantCalendarService,
  TenantUserService,
  MicrosoftTenant,
  MicrosoftTenantUser,
  MicrosoftTenantStatus,
} from '@checkfirst/nestjs-outlook';
import { CreateTenantEventDto } from './dto/create-tenant-event.dto';
import { RegisterUserMappingDto } from './dto/lookup-user.dto';

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
    @InjectRepository(MicrosoftTenant)
    private readonly tenantRepository: Repository<MicrosoftTenant>,
    @InjectRepository(MicrosoftTenantUser)
    private readonly tenantUserRepository: Repository<MicrosoftTenantUser>,
  ) {}

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

    if (!tenantUser) {
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
