import {
  Controller,
  Get,
  Post,
  Body,
  Param,
  Query,
  HttpCode,
  HttpStatus,
} from '@nestjs/common';
import { ApiTags, ApiOperation, ApiResponse, ApiParam } from '@nestjs/swagger';
import { TenantService } from './tenant.service';
import { CreateTenantEventDto } from './dto/create-tenant-event.dto';
import { LookupUserDto, RegisterUserMappingDto } from './dto/lookup-user.dto';
import { RegisterTenantDto } from './dto/register-tenant.dto';
import { GenerateCertificateDto } from './dto/generate-certificate.dto';

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
   * Generate a self-signed certificate for app-only authentication.
   *
   * Use `shared: true` for the one-time Checkfirst shared-app certificate, or
   * pass a `tenantId` to generate a dedicated per-tenant certificate. Returns
   * the public certificate PEM (to upload to Azure) and its thumbprint.
   */
  @Post('certificate/generate')
  @ApiOperation({
    summary: 'Generate a certificate',
    description:
      'Generates an RSA keypair + self-signed X.509 certificate and computes its x5t#S256 ' +
      'thumbprint. For the dedicated model the returned certificate PEM must be uploaded to ' +
      'the tenant\'s own Azure app registration. DEMO ONLY: keys are stored unencrypted on disk.',
  })
  @ApiResponse({ status: 201, description: 'Certificate generated' })
  generateCertificate(@Body() dto: GenerateCertificateDto) {
    return this.tenantService.generateCertificate(dto);
  }

  /**
   * Register a Microsoft 365 tenant for app-only access.
   *
   * This must be done before an administrator runs the admin-consent flow.
   * The response includes the admin-consent URL to hand to the tenant admin.
   */
  @Post('register')
  @ApiOperation({
    summary: 'Register a tenant',
    description:
      'Creates (or updates) the tenant record used for app-only authentication and returns ' +
      'the admin-consent URL. An Azure AD administrator must visit that URL to grant ' +
      'tenant-wide permissions before tenant calendar/user operations will work.',
  })
  @ApiResponse({ status: 201, description: 'Tenant registered; admin-consent URL returned' })
  async registerTenant(@Body() dto: RegisterTenantDto) {
    return this.tenantService.registerTenant(dto);
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
   * List users in the active tenant.
   */
  @Get('users')
  @ApiOperation({
    summary: 'List tenant users',
    description: 'Lists users in the active tenant via app-only auth. Optional OData $filter.',
  })
  @ApiResponse({ status: 200, description: 'Users listed successfully' })
  @ApiResponse({ status: 404, description: 'No active tenant configured' })
  async listUsers(@Query('filter') filter?: string) {
    return this.tenantService.listUsers(filter);
  }

  /**
   * Get a single user in the active tenant by UPN/email or object ID.
   */
  @Get('users/:userId')
  @ApiOperation({
    summary: 'Get tenant user',
    description: 'Gets a single user in the active tenant by UPN/email or object ID.',
  })
  @ApiParam({ name: 'userId', description: 'Microsoft UPN/email or object ID' })
  @ApiResponse({ status: 200, description: 'User found' })
  @ApiResponse({ status: 404, description: 'User not found' })
  async getUser(@Param('userId') userId: string) {
    return this.tenantService.getUser(userId);
  }

  /**
   * List a user's calendars in the active tenant.
   */
  @Get('users/:userId/calendars')
  @ApiOperation({
    summary: 'List user calendars',
    description: "Lists the specified user's calendars in the active tenant via app-only auth.",
  })
  @ApiParam({ name: 'userId', description: 'Microsoft UPN/email or object ID' })
  @ApiResponse({ status: 200, description: 'Calendars listed successfully' })
  async getUserCalendars(@Param('userId') userId: string) {
    return this.tenantService.getUserCalendars(userId);
  }

  /**
   * List a user's calendar events in the active tenant.
   */
  @Get('users/:userId/events')
  @ApiOperation({
    summary: 'List user events',
    description: "Lists the specified user's calendar events in the active tenant via app-only auth.",
  })
  @ApiParam({ name: 'userId', description: 'Microsoft UPN/email or object ID' })
  @ApiResponse({ status: 200, description: 'Events listed successfully' })
  async getUserEvents(@Param('userId') userId: string) {
    return this.tenantService.getUserEvents(userId);
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
