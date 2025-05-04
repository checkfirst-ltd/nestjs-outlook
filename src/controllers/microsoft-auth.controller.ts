import {
  Controller,
  Get,
  Query,
  Res,
  HttpStatus,
  BadRequestException,
  InternalServerErrorException,
} from '@nestjs/common';
import { Response } from 'express';
import { ApiTags, ApiResponse, ApiQuery, ApiOperation, ApiProduces } from '@nestjs/swagger';
import { MicrosoftAuthService } from '../services/microsoft-auth.service';

@ApiTags('Microsoft Auth')
@Controller('auth/microsoft')
export class MicrosoftAuthController {
  constructor(private readonly microsoftAuthService: MicrosoftAuthService) {}

  /**
   * OAuth callback endpoint for Microsoft authentication
   *
   * @summary Process Microsoft OAuth callback
   * @description This endpoint handles the OAuth callback from Microsoft after a user has
   * authenticated with their Microsoft account. It exchanges the authorization code for
   * access and refresh tokens, saves them for the user, and sets up necessary webhooks
   * for calendar synchronization.
   *
   * The user ID is extracted from the state parameter that was passed during the initial
   * authorization request. The state parameter is base64 encoded and contains user ID and CSRF token.
   *
   * @param {string} code - The authorization code from Microsoft
   * @param {string} state - Base64 encoded state containing user ID and CSRF token
   * @returns {HTML} HTML page confirming successful authentication
   *
   * @throws {BadRequestException} When code or state is missing or invalid
   * @throws {InternalServerErrorException} When authentication fails
   */
  @Get('callback')
  @ApiOperation({
    summary: 'Microsoft OAuth callback handler',
    description:
      'Processes the callback from Microsoft OAuth authentication flow. Exchanges the authorization code for access and refresh tokens, saves them for the user, and sets up necessary webhooks for calendar synchronization. The user ID is extracted from the state parameter.',
  })
  @ApiQuery({
    name: 'code',
    description: 'Authorization code from Microsoft',
    required: true,
    type: String,
    example: 'M.R3_BAY.c0def4c2-12bf-0b29-9a3a-f8a1c4f56789',
  })
  @ApiQuery({
    name: 'state',
    description: 'Base64 encoded state containing user ID and CSRF token',
    required: true,
    type: String,
    example: 'eyJ1c2VySWQiOiI3IiwiY3NyZiI6IjEyMzQ1In0',
  })
  @ApiResponse({
    status: 200,
    description: 'Authentication successful, tokens saved and webhooks created',
    content: {
      'text/html': {
        example:
          '<h1>Authorization successful!</h1><p>Your Microsoft account has been linked successfully.</p>',
      },
    },
  })
  @ApiResponse({
    status: 400,
    description: 'Invalid or missing code/state parameters',
  })
  @ApiResponse({
    status: 500,
    description: 'Server error during authentication process',
  })
  @ApiProduces('text/html')
  async callback(@Query('code') code: string, @Query('state') state: string, @Res() res: Response) {
    if (!code) {
      throw new BadRequestException('No code provided');
    }

    if (!state) {
      throw new BadRequestException('No state parameter provided');
    }

    try {
      await this.microsoftAuthService.exchangeCodeForToken(code, state);

      // Return success message HTML
      return res.status(HttpStatus.OK).send(`
        <h1>Authorization successful!</h1>
        <p>Your Microsoft account has been linked successfully.</p>
        <p>You can close this tab now.</p>
        <script>
          // Optionally notify the parent window
          if (window.opener) {
            window.opener.postMessage('microsoft-auth-success', '*');
          }
        </script>
      `);
    } catch (error) {
      console.error('Authentication failed:', error);
      throw new InternalServerErrorException('Authentication failed. Please try again.');
    }
  }
}
