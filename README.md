# NestJS Outlook

<p align="center">
  <a href="https://checkfirst.ai" target="_blank">
    <img src="https://raw.githubusercontent.com/checkfirst-ltd/nestjs-outlook/main/assets/checkfirst-logo.png" width="400" alt="CheckFirst Logo" />
  </a>
</p>

<p align="center">
  <a href="https://www.npmjs.com/package/@checkfirst/nestjs-outlook"><img src="https://img.shields.io/npm/v/@checkfirst/nestjs-outlook.svg" alt="NPM Version" /></a>
  <a href="https://www.npmjs.com/package/@checkfirst/nestjs-outlook"><img src="https://img.shields.io/npm/dm/@checkfirst/nestjs-outlook.svg" alt="NPM Downloads" /></a>
  <a href="https://github.com/checkfirst-ltd/nestjs-outlook/blob/main/LICENSE"><img src="https://img.shields.io/github/license/checkfirst-ltd/nestjs-outlook" alt="License" /></a>
</p>

An opinionated NestJS module for Microsoft Outlook integration that provides easy access to Microsoft Graph API for emails, calendars, and more.

## Features

- 🔄 Simplified Microsoft OAuth flow
- 📅 Calendar events management
- 🔔 Real-time webhooks for changes
- 🔐 Secure token storage and refresh

## Installation

```bash
npm install @checkfirst/nestjs-outlook
```

## Setup

### 1. Database Setup

This library requires two database tables in your application's database. Create these tables using a migration:

```typescript
import { MigrationInterface, QueryRunner } from 'typeorm';

export class CreateOutlookTables1697025846000 implements MigrationInterface {
  public async up(queryRunner: QueryRunner): Promise<void> {
    // Create outlook_webhook_subscriptions table
    await queryRunner.query(`
      CREATE TABLE outlook_webhook_subscriptions (
        id INTEGER PRIMARY KEY GENERATED ALWAYS AS IDENTITY,
        subscription_id VARCHAR(255) NOT NULL,
        user_id INTEGER NOT NULL,
        resource VARCHAR(255) NOT NULL,
        change_type VARCHAR(255) NOT NULL,
        client_state VARCHAR(255) NOT NULL,
        notification_url VARCHAR(255) NOT NULL,
        expiration_date_time TIMESTAMP NOT NULL,
        is_active BOOLEAN DEFAULT true,
        access_token TEXT,
        refresh_token TEXT,
        created_at TIMESTAMP DEFAULT NOW() NOT NULL,
        updated_at TIMESTAMP DEFAULT NOW() NOT NULL
      );
    `);

    // Create microsoft_csrf_tokens table
    await queryRunner.query(`
      CREATE TABLE microsoft_csrf_tokens (
        id INTEGER PRIMARY KEY GENERATED ALWAYS AS IDENTITY,
        token VARCHAR(64) NOT NULL,
        user_id VARCHAR(255) NOT NULL,
        expires TIMESTAMP NOT NULL,
        created_at TIMESTAMP DEFAULT NOW() NOT NULL,
        CONSTRAINT "UQ_microsoft_csrf_tokens_token" UNIQUE (token)
      );
    `);
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.query(`DROP TABLE IF EXISTS outlook_webhook_subscriptions`);
    await queryRunner.query(`DROP TABLE IF EXISTS microsoft_csrf_tokens`);
  }
}
```

You can customize this migration to match your database dialect (PostgreSQL, MySQL, etc.) if needed.

### 2. Import Required Modules

Register the module in your NestJS application and include the module entities in TypeORM:

```typescript
import { Module } from '@nestjs/common';
import { TypeOrmModule } from '@nestjs/typeorm';
import { EventEmitterModule } from '@nestjs/event-emitter';
import { ScheduleModule } from '@nestjs/schedule';
import { MicrosoftOutlookModule } from '@checkfirst/nestjs-outlook';
import * as path from 'path';

// Resolve the path to the outlook package
const outlookPackagePath = path.dirname(require.resolve('@checkfirst/nestjs-outlook/package.json'));

@Module({
  imports: [
    // Required modules
    TypeOrmModule.forRoot({
      // Your TypeORM configuration
      entities: [
        // Your app entities
        __dirname + '/**/*.entity{.ts,.js}',
        // Include outlook module entities
        path.join(outlookPackagePath, 'dist', 'entities', '*.entity.js')
      ],
    }),
    ScheduleModule.forRoot(),
    EventEmitterModule.forRoot(),
    
    // Microsoft Outlook Module
    MicrosoftOutlookModule.forRoot({
      clientId: 'YOUR_MICROSOFT_APP_CLIENT_ID',
      clientSecret: 'YOUR_MICROSOFT_APP_CLIENT_SECRET',
      redirectPath: 'auth/microsoft/callback',
      backendBaseUrl: 'https://your-api.example.com',
      basePath: 'api/v1',
    }),
  ],
})
export class AppModule {}
```

### 3. Create a Login Controller

Create a controller to handle Microsoft authentication:

```typescript
export class CalendarController {
  constructor(
    private readonly microsoftAuthService: MicrosoftAuthService
  ) {}

  @UseGuards(AuthGuard)
  @Get('auth/microsoft/login')
  async login(@Req() req: Request) {
    const user = req.user as UserTokenPayload;
    if (!user?.id) {
      throw new ForbiddenException('User not authenticated');
    }
        
    // Get the login URL from the Microsoft auth service
    return await this.microsoftAuthService.getLoginUrl(user.id.toString());
  }
}
```

## Available Services

- `OutlookService` - Main service for Microsoft Graph API operations on Outlook
- `MicrosoftAuthService` - For authentication and token management

## Events

The library uses NestJS's EventEmitter to emit events for various Outlook activities. You can listen to these events in your application to react to changes.

### Available Events

The library exposes event types through the `OutlookEventTypes` enum:

- `OutlookEventTypes.AUTH_TOKENS_SAVE` - Emitted when OAuth tokens are initially saved
- `OutlookEventTypes.AUTH_TOKENS_UPDATE` - Emitted when OAuth tokens are refreshed
- `OutlookEventTypes.EVENT_CREATED` - Emitted when a new Outlook calendar event is created via webhook
- `OutlookEventTypes.EVENT_UPDATED` - Emitted when an Outlook calendar event is updated via webhook
- `OutlookEventTypes.EVENT_DELETED` - Emitted when an Outlook calendar event is deleted via webhook

### Listening to Events

You can listen to these events in your application using the `@OnEvent` decorator from `@nestjs/event-emitter` and the `OutlookEventTypes` enum:

```typescript
import { Injectable } from '@nestjs/common';
import { OnEvent } from '@nestjs/event-emitter';
import { OutlookEventTypes, OutlookResourceData } from '@checkfirst/nestjs-outlook';

@Injectable()
export class YourService {
  @OnEvent(OutlookEventTypes.EVENT_CREATED)
  handleOutlookEventCreated(data: OutlookResourceData) {
    console.log('New Outlook event created:', data.id);
    // Handle the new event...
  }

  @OnEvent(OutlookEventTypes.EVENT_UPDATED)
  handleOutlookEventUpdated(data: OutlookResourceData) {
    console.log('Outlook event updated:', data.id);
    // Handle the updated event...
  }

  @OnEvent(OutlookEventTypes.EVENT_DELETED)
  handleOutlookEventDeleted(data: OutlookResourceData) {
    console.log('Outlook event deleted:', data.id);
    // Handle the deleted event...
  }
}
```

These events are emitted when Microsoft Graph sends webhook notifications to your application for calendar events changes.

## Support

- [GitHub Issues](https://github.com/checkfirst-ltd/nestjs-outlook/issues)
- [Documentation](https://github.com/checkfirst-ltd/nestjs-outlook#readme)

## Contributing

We welcome contributions! Please see our [Contributing Guide](CONTRIBUTING.md) for more details.

## Code of Conduct

This project adheres to a Code of Conduct that all participants are expected to follow. Please read the [Code of Conduct](https://github.com/checkfirst-ltd/nestjs-outlook/blob/main/CONTRIBUTING.md#code-of-conduct) for details on our expectations.

## About CheckFirst

<a href="https://checkfirst.ai" target="_blank">
    <img src="https://raw.githubusercontent.com/checkfirst-ltd/nestjs-outlook/main/assets/checkfirst-logo.png" width="400" alt="CheckFirst Logo" />
</a>

[CheckFirst](https://checkfirst.ai) is a trusted provider of developer tools and solutions. We build open-source libraries that help developers create better applications faster.

## License

[MIT](LICENSE) © [CheckFirst Ltd](https://checkfirst.ai) 