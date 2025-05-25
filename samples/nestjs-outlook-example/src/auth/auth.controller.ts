import { Controller, Get } from '@nestjs/common';
import { MicrosoftAuthService, PermissionScope } from '@checkfirst/nestjs-outlook';

@Controller('auth')
export class AuthController {
  constructor(
    private readonly microsoftAuthService: MicrosoftAuthService,
  ) {}

  @Get('microsoft/login')
  async login() {
    // In a real application, you would have proper authentication
    // For this example, we're using a mock user ID
    const mockUserId = '1';
    
    // Pass the permission scopes to the login URL. They are defined in PermissionScope enum
    return await this.microsoftAuthService.getLoginUrl(mockUserId, [
      PermissionScope.CALENDAR_READ,
      PermissionScope.CALENDAR_WRITE,
      PermissionScope.EMAIL_READ,
      PermissionScope.EMAIL_WRITE,
      PermissionScope.EMAIL_SEND,
    ]);
  }
} 