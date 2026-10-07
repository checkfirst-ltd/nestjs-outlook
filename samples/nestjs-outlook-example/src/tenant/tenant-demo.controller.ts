import {
  Controller,
  Get,
  Post,
  Body,
  Param,
  Query,
  HttpException,
  HttpStatus,
} from '@nestjs/common';
import {
  ApiTags,
  ApiOperation,
  ApiResponse,
  ApiParam,
  ApiQuery,
} from '@nestjs/swagger';
import {
  TenantCalendarService,
  TenantUserService,
  TenantUserLookupResult,
} from '@checkfirst/nestjs-outlook';
import { LookupUserDto, RegisterUserMappingDto } from './dto/lookup-user.dto';
import { CreateTenantEventDto } from './dto/create-tenant-event.dto';

/**
 * Demo controller for tenant-wide Microsoft 365 operations using app-only authentication.
 *
 * These endpoints demonstrate how to use TenantCalendarService and TenantUserService
 * to access calendars for any user within a Microsoft 365 tenant without per-user OAuth.
 *
 * Prerequisites:
 * - A registered Microsoft tenant with admin consent
 * - Certificate-based app-only authentication configured
 * - User mappings registered via the registerUserMapping endpoint
 */
@ApiTags('Tenant Demo')
@Controller('tenant-demo')
export class TenantDemoController {
  /**
   * Demo tenant ID - in production, retrieve this from your database
   * based on the authenticated user's organization
   */
  private readonly DEMO_TENANT_ID = process.env.DEMO_TENANT_ID || 'your-tenant-id';

  constructor(
    private readonly tenantCalendarService: TenantCalendarService,
    private readonly tenantUserService: TenantUserService,
  ) {}

  // ─────────────────────────────────────────────────────────────────────────────
  // User Lookup Endpoints
  // ─────────────────────────────────────────────────────────────────────────────

  @Get('users/lookup')
  @ApiOperation({
    summary: 'Look up a Microsoft user by email',
    description: 'Searches for a user in the Microsoft tenant by their email address or UPN',
  })
  @ApiQuery({ name: 'email', required: true, description: 'Email address to look up' })
  @ApiResponse({
    status: 200,
    description: 'User found',
    schema: {
      type: 'object',
      properties: {
        microsoftUserId: { type: 'string' },
        userPrincipalName: { type: 'string' },
        displayName: { type: 'string' },
        email: { type: 'string', nullable: true },
      },
    },
  })
  @ApiResponse({ status: 404, description: 'User not found in tenant' })
  async lookupUser(@Query('email') email: string): Promise<TenantUserLookupResult> {
    if (!email) {
      throw new HttpException('Email query parameter is required', HttpStatus.BAD_REQUEST);
    }

    const user = await this.tenantUserService.lookupUserByEmail(this.DEMO_TENANT_ID, email);

    if (!user) {
      throw new HttpException(`User not found: ${email}`, HttpStatus.NOT_FOUND);
    }

    return user;
  }

  @Get('users/:microsoftUserId')
  @ApiOperation({
    summary: 'Get user details by Microsoft user ID',
    description: 'Retrieves user information from Microsoft Graph using the user ID',
  })
  @ApiParam({ name: 'microsoftUserId', description: 'Microsoft Graph user ID (GUID)' })
  @ApiResponse({ status: 200, description: 'User details retrieved' })
  @ApiResponse({ status: 404, description: 'User not found' })
  async getUserById(
    @Param('microsoftUserId') microsoftUserId: string,
  ): Promise<TenantUserLookupResult> {
    const user = await this.tenantUserService.getUserById(this.DEMO_TENANT_ID, microsoftUserId);

    if (!user) {
      throw new HttpException(`User not found: ${microsoftUserId}`, HttpStatus.NOT_FOUND);
    }

    return user;
  }

  @Get('users')
  @ApiOperation({
    summary: 'List users in the tenant',
    description: 'Returns a paginated list of users in the Microsoft tenant',
  })
  @ApiQuery({ name: 'top', required: false, description: 'Max users to return (default: 10)' })
  @ApiQuery({ name: 'filter', required: false, description: 'OData filter (e.g., accountEnabled eq true)' })
  @ApiResponse({ status: 200, description: 'List of users' })
  async listUsers(
    @Query('top') top?: string,
    @Query('filter') filter?: string,
  ) {
    const result = await this.tenantUserService.listUsers(this.DEMO_TENANT_ID, {
      top: top ? parseInt(top, 10) : 10,
      filter,
    });

    return {
      users: result.users,
      hasMore: !!result.nextLink,
    };
  }

  @Post('users/register')
  @ApiOperation({
    summary: 'Register a user mapping',
    description: 'Maps an external user ID from your application to a Microsoft user',
  })
  @ApiResponse({ status: 201, description: 'User mapping created' })
  @ApiResponse({ status: 404, description: 'Microsoft user not found' })
  async registerUserMapping(@Body() dto: RegisterUserMappingDto) {
    try {
      const mapping = await this.tenantUserService.registerUserMapping(
        this.DEMO_TENANT_ID,
        dto.externalUserId,
        dto.email,
      );

      return {
        success: true,
        externalUserId: mapping.externalUserId,
        microsoftUserId: mapping.microsoftUserId,
        userPrincipalName: mapping.userPrincipalName,
      };
    } catch (error) {
      if (error instanceof Error && error.message.includes('not found')) {
        throw new HttpException(error.message, HttpStatus.NOT_FOUND);
      }
      throw error;
    }
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // Calendar Endpoints
  // ─────────────────────────────────────────────────────────────────────────────

  @Get('calendar/:externalUserId')
  @ApiOperation({
    summary: 'Get default calendar ID for a user',
    description: 'Retrieves the default calendar ID for a mapped user',
  })
  @ApiParam({ name: 'externalUserId', description: 'Your application user ID' })
  @ApiResponse({ status: 200, description: 'Calendar ID retrieved' })
  @ApiResponse({ status: 404, description: 'User mapping not found' })
  async getDefaultCalendarId(@Param('externalUserId') externalUserId: string) {
    // First, get the Microsoft user ID from the mapping
    const microsoftUserId = await this.tenantUserService.getMicrosoftUserId(
      this.DEMO_TENANT_ID,
      externalUserId,
    );

    if (!microsoftUserId) {
      throw new HttpException(
        `No user mapping found for external ID: ${externalUserId}. Register the user first.`,
        HttpStatus.NOT_FOUND,
      );
    }

    const calendarId = await this.tenantCalendarService.getDefaultCalendarId(
      this.DEMO_TENANT_ID,
      microsoftUserId,
    );

    return {
      externalUserId,
      microsoftUserId,
      calendarId,
    };
  }

  @Post('calendar/events')
  @ApiOperation({
    summary: 'Create a calendar event for a user',
    description: 'Creates an event in the mapped user\'s calendar using app-only auth',
  })
  @ApiResponse({ status: 201, description: 'Event created' })
  @ApiResponse({ status: 404, description: 'User mapping not found' })
  async createEvent(@Body() dto: CreateTenantEventDto) {
    // Get the Microsoft user ID from the mapping
    const microsoftUserId = await this.tenantUserService.getMicrosoftUserId(
      this.DEMO_TENANT_ID,
      dto.externalUserId,
    );

    if (!microsoftUserId) {
      throw new HttpException(
        `No user mapping found for external ID: ${dto.externalUserId}. Register the user first.`,
        HttpStatus.NOT_FOUND,
      );
    }

    // Get the user's default calendar
    const calendarId = await this.tenantCalendarService.getDefaultCalendarId(
      this.DEMO_TENANT_ID,
      microsoftUserId,
    );

    // Build the event object for Microsoft Graph
    const event = {
      subject: dto.subject,
      start: {
        dateTime: dto.startDateTime,
        timeZone: 'UTC',
      },
      end: {
        dateTime: dto.endDateTime,
        timeZone: 'UTC',
      },
      body: dto.body
        ? {
            contentType: 'html' as const,
            content: dto.body,
          }
        : undefined,
      location: dto.location
        ? {
            displayName: dto.location,
          }
        : undefined,
    };

    const result = await this.tenantCalendarService.createEvent(
      event,
      this.DEMO_TENANT_ID,
      microsoftUserId,
      calendarId,
    );

    return {
      success: true,
      eventId: result.event.id,
      webLink: result.event.webLink,
      subject: result.event.subject,
    };
  }

  @Get('calendar/events/:externalUserId/:eventId')
  @ApiOperation({
    summary: 'Get a calendar event by ID',
    description: 'Retrieves event details by event ID for a mapped user',
  })
  @ApiParam({ name: 'externalUserId', description: 'Your application user ID' })
  @ApiParam({ name: 'eventId', description: 'Microsoft Graph event ID' })
  @ApiResponse({ status: 200, description: 'Event retrieved' })
  @ApiResponse({ status: 404, description: 'Event or user not found' })
  async getEvent(
    @Param('externalUserId') externalUserId: string,
    @Param('eventId') eventId: string,
  ) {
    const microsoftUserId = await this.tenantUserService.getMicrosoftUserId(
      this.DEMO_TENANT_ID,
      externalUserId,
    );

    if (!microsoftUserId) {
      throw new HttpException(
        `No user mapping found for external ID: ${externalUserId}`,
        HttpStatus.NOT_FOUND,
      );
    }

    const event = await this.tenantCalendarService.getEventById(
      this.DEMO_TENANT_ID,
      microsoftUserId,
      eventId,
    );

    if (!event) {
      throw new HttpException(`Event not found: ${eventId}`, HttpStatus.NOT_FOUND);
    }

    return event;
  }

  @Get('calendar/events/:externalUserId')
  @ApiOperation({
    summary: 'List calendar events for a user',
    description: 'Streams calendar events for a mapped user within a date range',
  })
  @ApiParam({ name: 'externalUserId', description: 'Your application user ID' })
  @ApiQuery({ name: 'startDate', required: false, description: 'Start date (ISO 8601)' })
  @ApiQuery({ name: 'endDate', required: false, description: 'End date (ISO 8601)' })
  @ApiQuery({ name: 'limit', required: false, description: 'Max events to return (default: 50)' })
  @ApiResponse({ status: 200, description: 'List of events' })
  async listEvents(
    @Param('externalUserId') externalUserId: string,
    @Query('startDate') startDate?: string,
    @Query('endDate') endDate?: string,
    @Query('limit') limit?: string,
  ) {
    const microsoftUserId = await this.tenantUserService.getMicrosoftUserId(
      this.DEMO_TENANT_ID,
      externalUserId,
    );

    if (!microsoftUserId) {
      throw new HttpException(
        `No user mapping found for external ID: ${externalUserId}`,
        HttpStatus.NOT_FOUND,
      );
    }

    const maxEvents = limit ? parseInt(limit, 10) : 50;
    const events: unknown[] = [];

    // Use the streaming API but collect into an array for demo purposes
    const stream = this.tenantCalendarService.streamEvents(this.DEMO_TENANT_ID, microsoftUserId, {
      startDate: startDate ? new Date(startDate) : undefined,
      endDate: endDate ? new Date(endDate) : undefined,
      batchSize: Math.min(maxEvents, 100),
    });

    for await (const batch of stream) {
      events.push(...batch);
      if (events.length >= maxEvents) {
        break;
      }
    }

    return {
      externalUserId,
      microsoftUserId,
      count: Math.min(events.length, maxEvents),
      events: events.slice(0, maxEvents),
    };
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // Health Check
  // ─────────────────────────────────────────────────────────────────────────────

  @Get('health')
  @ApiOperation({
    summary: 'Check tenant services health',
    description: 'Verifies that tenant services are properly configured',
  })
  @ApiResponse({ status: 200, description: 'Services are healthy' })
  healthCheck() {
    return {
      status: 'ok',
      tenantId: this.DEMO_TENANT_ID,
      services: {
        tenantCalendarService: !!this.tenantCalendarService,
        tenantUserService: !!this.tenantUserService,
      },
      timestamp: new Date().toISOString(),
    };
  }
}
