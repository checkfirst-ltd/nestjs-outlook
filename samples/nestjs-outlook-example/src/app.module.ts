import { Module } from '@nestjs/common';
import { TypeOrmModule } from '@nestjs/typeorm';
import { EventEmitterModule } from '@nestjs/event-emitter';
import { ScheduleModule } from '@nestjs/schedule';
import { ConfigModule } from '@nestjs/config';
import { CalendarModule } from './calendar/calendar.module';
import * as path from 'path';
// Resolve the path to the nestjs-outlook package - works with npm link
const outlookPackagePath = path.dirname(require.resolve('@checkfirst/nestjs-outlook/package.json'));

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
    TypeOrmModule.forRoot({
      type: 'sqlite',
      database: 'db.sqlite',
      entities: [
        __dirname + '/**/*.entity{.ts,.js}',
        path.join(outlookPackagePath, 'dist', 'entities', '*.entity.js'),
      ],
      synchronize: true, // Don't use this in production
    }),
    ScheduleModule.forRoot(),
    EventEmitterModule.forRoot(),
    CalendarModule,
  ],
})
export class AppModule {} 