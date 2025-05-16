import { Module } from '@nestjs/common';
import { CalendarController } from './calendar.controller';
import { CalendarService } from './calendar.service';
import { TypeOrmModule } from '@nestjs/typeorm';
import { UserCalendar } from './entities/user-calendar.entity';
import { UserCalendarRepository } from './repositories/user-calendar.repository';
import { User } from '../users/entities/user.entity';
import { MicrosoftOutlookModule } from '@checkfirst/nestjs-outlook';
import { ConfigModule, ConfigService } from '@nestjs/config';

@Module({
  imports: [
    TypeOrmModule.forFeature([UserCalendar, User]),
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
  controllers: [CalendarController],
  providers: [
    CalendarService, 
    UserCalendarRepository,
  ],
  exports: [CalendarService, UserCalendarRepository],
})
export class CalendarModule {} 