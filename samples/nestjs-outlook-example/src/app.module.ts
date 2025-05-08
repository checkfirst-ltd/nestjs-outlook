import { Module } from '@nestjs/common';
import { ConfigModule, ConfigService } from '@nestjs/config';
import { TypeOrmModule } from '@nestjs/typeorm';
import { EventEmitterModule } from '@nestjs/event-emitter';
import { ScheduleModule } from '@nestjs/schedule';
import { CalendarModule } from './calendar/calendar.module';
import { EmailModule } from './email/email.module';
import { AuthModule } from './auth/auth.module';
import { typeOrmModuleOptions } from './config/database.config';

@Module({
  imports: [
    ConfigModule.forRoot({
      isGlobal: true,
      envFilePath: ['.env.development.local', '.env.development', '.env'],
      cache: true,
      expandVariables: true,
      // Log loaded environment variables during startup
      load: [() => {
        console.log('Loaded environment variables:', {
          MICROSOFT_CLIENT_ID: process.env.MICROSOFT_CLIENT_ID ? 'set' : 'not set',
          MICROSOFT_CLIENT_SECRET: process.env.MICROSOFT_CLIENT_SECRET ? 'set' : 'not set',
          BACKEND_BASE_URL: process.env.BACKEND_BASE_URL,
          MICROSOFT_REDIRECT_PATH: process.env.MICROSOFT_REDIRECT_PATH,
          MICROSOFT_BASE_PATH: process.env.MICROSOFT_BASE_PATH,
        });
        return {};
      }],
    }),
    TypeOrmModule.forRoot(typeOrmModuleOptions),
    ScheduleModule.forRoot(),
    EventEmitterModule.forRoot(),
    // Import our feature modules
    AuthModule,
    CalendarModule,
    EmailModule,
  ],
})
export class AppModule {} 