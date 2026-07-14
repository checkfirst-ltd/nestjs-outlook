import {
  Controller, Get, Post, Param, Query, Body, Res, HttpStatus, Logger, Optional, Inject, BadRequestException,
} from '@nestjs/common';
import { Response } from 'express';
import { ApiTags, ApiOperation, ApiParam, ApiQuery, ApiBody, ApiResponse } from '@nestjs/swagger';
import { HealthService, UserHealth } from '../services/health/health.service';
import { HealthCheckDto } from '../dto/health-check.dto';

/**
 * Endpoints to diagnose and recover user connection health (delegated + app-only).
 */
@ApiTags('Health')
@Controller('auth/microsoft/health')
export class HealthController {
  private readonly logger = new Logger(HealthController.name);

  constructor(
    @Optional()
    @Inject(HealthService)
    private readonly healthService: HealthService | null,
  ) {}

  /**
   * Diagnose a single user's connection health.
   */
  @Get(':externalUserId')
  @ApiOperation({
    summary: 'Get one user\'s connection health',
    description:
      'Returns a health verdict combining the microsoft_users row and its active calendar ' +
      'subscription. Pass verifyAtGraph=true to also confirm the subscription exists at Microsoft.',
  })
  @ApiParam({ name: 'externalUserId', description: "The host application's user id", example: 'insp-001' })
  @ApiQuery({ name: 'verifyAtGraph', required: false, type: Boolean })
  @ApiResponse({ status: 200, description: 'Health verdict for the user' })
  async getUserHealth(
    @Param('externalUserId') externalUserId: string,
    @Query('verifyAtGraph') verifyAtGraph?: string,
  ): Promise<UserHealth> {
    const service = this.requireService();
    return service.checkUser(externalUserId, { verifyAtGraph: this.isTruthy(verifyAtGraph) });
  }

  /**
   * Diagnose a batch of users (read-only).
   */
  @Post('check')
  @ApiOperation({
    summary: 'Check connection health for many users',
    description: 'Read-only bulk health verdicts. Set verifyAtGraph to also confirm each at Microsoft.',
  })
  @ApiBody({ type: HealthCheckDto })
  @ApiResponse({ status: 200, description: 'Health verdicts (input order preserved)' })
  async checkUsers(@Body() body: HealthCheckDto): Promise<UserHealth[]> {
    const service = this.requireService();
    const externalUserIds = this.validateExternalUserIds(body);
    return service.checkUsers(externalUserIds, { verifyAtGraph: this.isTruthy(body.verifyAtGraph) });
  }

  /**
   * Diagnose and recover a batch of users. Recreates fixable subscriptions (delegated or app-only)
   * and reports the rest. Runs in the background and returns 202; listen for the
   * `outlook.user.health.recovery.completed` event for the summary.
   */
  @Post('recover')
  @ApiOperation({
    summary: 'Check and recover many users',
    description:
      'Auto-fixes recoverable users (recreates missing/expired/stale/gone-at-Graph subscriptions) ' +
      'and reports those needing a human. Runs in the background (202); listen for the ' +
      "'outlook.user.health.recovery.completed' event for the result.",
  })
  @ApiBody({ type: HealthCheckDto })
  @ApiResponse({ status: 202, description: 'Recovery accepted and running in the background' })
  @ApiResponse({ status: 400, description: 'Invalid body' })
  recoverUsers(
    @Body() body: HealthCheckDto,
    @Res({ passthrough: true }) res?: Response,
  ): { message: string; totalRequested: number } {
    const service = this.requireService();
    const externalUserIds = this.validateExternalUserIds(body);
    const verifyAtGraph = this.isTruthy(body.verifyAtGraph);

    // Recovery does per-user Graph work — run detached so a large batch can't time out the request.
    service.recoverUsers(externalUserIds, { verifyAtGraph }).catch((error: unknown) => {
      this.logger.error(
        `[recoverUsers] Background health recovery failed: ${
          error instanceof Error ? error.message : 'Unknown error'
        }`,
      );
    });

    if (res) {
      res.status(HttpStatus.ACCEPTED);
    }
    return {
      message:
        'Health recovery is running in the background. Listen for the ' +
        "'outlook.user.health.recovery.completed' event for the result.",
      totalRequested: externalUserIds.length,
    };
  }

  private requireService(): HealthService {
    if (!this.healthService) {
      throw new Error('Health service is not available in this application');
    }
    return this.healthService;
  }

  /** Validate the request body's user id list (no class-transformer, so validated here). */
  private validateExternalUserIds(body: HealthCheckDto): string[] {
    const raw: unknown = body.externalUserIds;
    if (!Array.isArray(raw) || raw.length === 0) {
      throw new BadRequestException('`externalUserIds` must be a non-empty array');
    }
    for (const id of raw) {
      if (typeof id !== 'string' || id.trim() === '') {
        throw new BadRequestException('each externalUserId must be a non-empty string');
      }
    }
    return raw as string[];
  }

  private isTruthy(value: string | boolean | undefined): boolean {
    if (typeof value === 'boolean') {
      return value;
    }
    if (typeof value === 'string') {
      const v = value.trim().toLowerCase();
      return v === 'true' || v === '1' || v === 'yes';
    }
    return false;
  }
}
