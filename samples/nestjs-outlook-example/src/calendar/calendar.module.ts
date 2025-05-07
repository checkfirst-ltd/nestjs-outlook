import { Module } from '@nestjs/common';
import { MicrosoftOutlookModule } from '@checkfirst/nestjs-outlook';
import { CalendarController } from './calendar.controller';
import { CalendarService } from './calendar.service';
import { ConfigModule, ConfigService } from '@nestjs/config';
import { TypeOrmModule } from '@nestjs/typeorm';
import { UserCalendar } from './entities/user-calendar.entity';
import { UserCalendarRepository } from './repositories/user-calendar.repository';
import { User } from '../users/entities/user.entity';

@Module({
  imports: [
    TypeOrmModule.forFeature([UserCalendar, User]),
    MicrosoftOutlookModule.forRootAsync({
      imports: [ConfigModule],
      inject: [ConfigService],
      useFactory: (configService: ConfigService) => ({
        clientId: configService.get('MICROSOFT_CLIENT_ID'),
        clientSecret: configService.get('MICROSOFT_CLIENT_SECRET'),
        redirectPath: configService.get('MICROSOFT_REDIRECT_PATH', 'auth/microsoft/callback'),
        backendBaseUrl: configService.get('BACKEND_BASE_URL', 'http://localhost:3000'),
        basePath: configService.get('MICROSOFT_BASE_PATH', 'api/v1'),
      }),
    }),
  ],
  controllers: [CalendarController],
  providers: [CalendarService, UserCalendarRepository],
})
export class CalendarModule {} 