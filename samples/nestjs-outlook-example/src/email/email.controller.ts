import { Body, Controller, Post } from '@nestjs/common';
import { EmailService } from './email.service';
import { SendEmailDto } from './dto/send-email.dto';

@Controller('email')
export class EmailController {
  constructor(private readonly emailService: EmailService) {}

  @Post('send')
  async sendEmail(@Body() sendEmailDto: SendEmailDto) {
    // For this sample, hardcode the userId - in a real app, you'd get this from auth
    const userId = 1;
    
    return await this.emailService.sendEmail(
      userId,
      sendEmailDto.to,
      sendEmailDto.subject,
      sendEmailDto.body
    );
  }
} 