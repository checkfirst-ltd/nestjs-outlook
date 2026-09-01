import { Module } from '@nestjs/common';
import { ConfigModule } from '@nestjs/config';
import { TypeOrmModule } from '@nestjs/typeorm';
import { EventEmitterModule } from '@nestjs/event-emitter';
import { ScheduleModule } from '@nestjs/schedule';
import { ServeStaticModule } from '@nestjs/serve-static';
import { join } from 'path';
import { CalendarModule } from './calendar/calendar.module';
import { EmailModule } from './email/email.module';
import { AuthModule } from './auth/auth.module';
import { TenantModule } from './tenant/tenant.module';
import { TenantDemoModule } from './tenant/tenant-demo.module';
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
    // Serve static files from /public directory
    ServeStaticModule.forRoot({
      rootPath: join(__dirname, '..', 'public'),
      serveRoot: '/',
      exclude: ['/api*', '/auth*', '/calendar*', '/email*', '/tenant*', '/tenant-demo*'],
    }),
    // Import our feature modules
    AuthModule,
    CalendarModule,
    EmailModule,
    TenantModule,
    TenantDemoModule,
  ],
})
export class AppModule {} 