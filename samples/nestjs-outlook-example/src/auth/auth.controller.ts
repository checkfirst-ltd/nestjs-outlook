import { Controller, Get, Req } from '@nestjs/common';
import { MicrosoftAuthService } from '@checkfirst/nestjs-outlook';

@Controller('auth')
export class AuthController {
  constructor(
    private readonly microsoftAuthService: MicrosoftAuthService,
  ) {}

  @Get('microsoft/login')
  async login(@Req() req: any) {
    // In a real application, you would have proper authentication
    // For this example, we're using a mock user ID
    const mockUserId = '1';
    return await this.microsoftAuthService.getLoginUrl(mockUserId);
  }
} 