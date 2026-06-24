import {
  Controller,
  Get,
  Post,
  Body,
  Param,
  HttpCode,
  HttpStatus,
} from '@nestjs/common';
import { ApiTags, ApiOperation, ApiResponse, ApiParam } from '@nestjs/swagger';
import { TenantService } from './tenant.service';
import { CreateTenantEventDto } from './dto/create-tenant-event.dto';
import { LookupUserDto, RegisterUserMappingDto } from './dto/lookup-user.dto';

/**
 * Controller for tenant-wide calendar operations using app-only authentication.
 *
 * These endpoints demonstrate enterprise scenarios where the application
 * needs to access calendars across all users in a Microsoft 365 tenant
 * without requiring individual user consent.
 */
@ApiTags('Tenant Calendar')
@Controller('tenant')
export class TenantController {
  constructor(private readonly tenantService: TenantService) {}

  /**
   * Get tenant status and configuration summary.
   */
  @Get('status')
  @ApiOperation({
    summary: 'Get tenant status',
    description: 'Returns the current tenant configuration and status.',
  })
  @ApiResponse({ status: 200, description: 'Tenant status retrieved successfully' })
  async getTenantStatus() {
    return this.tenantService.getTenantStatus();
  }

  /**
   * Look up a Microsoft user by email address.
   */
  @Post('users/lookup')
  @HttpCode(HttpStatus.OK)
  @ApiOperation({
    summary: 'Look up user by email',
    description: 'Looks up a Microsoft user in the tenant by their email address.',
  })
  @ApiResponse({ status: 200, description: 'User found' })
  @ApiResponse({ status: 404, description: 'User not found' })
  async lookupUser(@Body() dto: LookupUserDto) {
    return this.tenantService.lookupUser(dto.email);
  }

  /**
   * Register a mapping between an external user ID and a Microsoft user.
   */
  @Post('users/register')
  @ApiOperation({
    summary: 'Register user mapping',
    description:
      'Maps an external user ID from your application to a Microsoft user in the tenant. ' +
      'This mapping is required before you can access the user\'s calendar.',
  })
  @ApiResponse({ status: 201, description: 'User mapping registered successfully' })
  @ApiResponse({ status: 404, description: 'Microsoft user not found' })
  async registerUserMapping(@Body() dto: RegisterUserMappingDto) {
    return this.tenantService.registerUserMapping(dto);
  }

  /**
   * Get the default calendar ID for a user.
   */
  @Get('users/:externalUserId/calendar')
  @ApiOperation({
    summary: 'Get user default calendar',
    description: 'Gets the default calendar ID for the specified user.',
  })
  @ApiParam({ name: 'externalUserId', description: 'External user ID from your application' })
  @ApiResponse({ status: 200, description: 'Calendar info retrieved successfully' })
  @ApiResponse({ status: 404, description: 'User mapping not found' })
  async getDefaultCalendar(@Param('externalUserId') externalUserId: string) {
    return this.tenantService.getDefaultCalendarId(externalUserId);
  }

  /**
   * Get a specific event by ID.
   */
  @Get('users/:externalUserId/events/:eventId')
  @ApiOperation({
    summary: 'Get calendar event',
    description: 'Retrieves a specific calendar event by ID.',
  })
  @ApiParam({ name: 'externalUserId', description: 'External user ID from your application' })
  @ApiParam({ name: 'eventId', description: 'Microsoft Graph event ID' })
  @ApiResponse({ status: 200, description: 'Event retrieved successfully' })
  @ApiResponse({ status: 404, description: 'Event or user not found' })
  async getEvent(
    @Param('externalUserId') externalUserId: string,
    @Param('eventId') eventId: string,
  ) {
    return this.tenantService.getEvent(externalUserId, eventId);
  }

  /**
   * Create a calendar event for a user.
   */
  @Post('events')
  @ApiOperation({
    summary: 'Create calendar event',
    description:
      'Creates a calendar event for the specified user. ' +
      'The user must have been registered via POST /tenant/users/register first.',
  })
  @ApiResponse({ status: 201, description: 'Event created successfully' })
  @ApiResponse({ status: 404, description: 'User mapping not found' })
  async createEvent(@Body() dto: CreateTenantEventDto) {
    return this.tenantService.createEvent(dto);
  }
}
