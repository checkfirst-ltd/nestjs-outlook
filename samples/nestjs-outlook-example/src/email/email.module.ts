import { Module } from '@nestjs/common';
import { TypeOrmModule } from '@nestjs/typeorm';
import { EmailController } from './email.controller';
import { EmailService } from './email.service';
import { User } from '../users/entities/user.entity';
import { UserCalendar } from '../calendar/entities/user-calendar.entity';
import { UserCalendarRepository } from '../calendar/repositories/user-calendar.repository';
import { MicrosoftOutlookModule } from '@checkfirst/nestjs-outlook';
import { ConfigModule, ConfigService } from '@nestjs/config';

@Module({
  imports: [
    TypeOrmModule.forFeature([User, UserCalendar]),
    // Initialize MicrosoftOutlookModule for this feature module
    MicrosoftOutlookModule.forRootAsync({
      imports: [ConfigModule],
      inject: [ConfigService],
      useFactory: (configService: ConfigService) => ({
        clientId: configService.get('MICROSOFT_CLIENT_ID'),
        clientSecret: configService.get('MICROSOFT_CLIENT_SECRET'),
        redirectPath: configService.get('MICROSOFT_REDIRECT_PATH', 'auth/microsoft/callback'),
        backendBaseUrl: configService.get('BACKEND_BASE_URL'),
        basePath: configService.get('MICROSOFT_BASE_PATH'),
      }),
    }),
  ],
  controllers: [EmailController],
  providers: [
    EmailService, 
    UserCalendarRepository,
  ],
})
export class EmailModule {} 