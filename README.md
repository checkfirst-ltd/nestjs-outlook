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

## Why We Built This

At Checkfirst, we believe that tools should be as transparent as the systems they support. This library is part of our commitment to the Testing, Inspection, Certification, and Compliance (TICC) industry, where we help transform operations through intelligent automation. While our [ScheduleAI platform](https://www.checkfirst.ai/scheduleai) helps organizations optimize inspections, audits and fieldwork, we recognize that true innovation requires open collaboration. By sharing the Microsoft integration layer that powers our authentication, calendar, and email services, we're enabling developers to build robust, enterprise-grade applications without reimplementing complex Microsoft Graph protocols. Whether you're creating scheduling systems, communication tools, or productivity enhancers, this library embodies our philosophy that trust in technology starts with transparency and accessibility.

<!-- START doctoc generated TOC please keep comment here to allow auto update -->
<!-- DON'T EDIT THIS SECTION, INSTEAD RE-RUN doctoc TO UPDATE -->
## Table of Contents

- [Features](#features)
- [Installation](#installation)
- [Setup](#setup)
  - [1. Database Setup](#1-database-setup)
  - [2. Microsoft App Registration](#2-microsoft-app-registration)
  - [3. Import Required Modules](#3-import-required-modules)
  - [4. Create an Auth Controller](#4-create-an-auth-controller)
- [Permission Scopes](#permission-scopes)
- [Available Services and Controllers](#available-services-and-controllers)
  - [1. MicrosoftAuthService and MicrosoftAuthController](#1-microsoftauthservice-and-microsoftauthcontroller)
  - [2. CalendarService and CalendarController](#2-calendarservice-and-calendarcontroller)
  - [3. EmailService](#3-emailservice)
- [Events](#events)
  - [Available Events](#available-events)
  - [Listening to Events](#listening-to-events)
- [Example Application Architecture](#example-application-architecture)
- [Local Development](#local-development)
  - [Prerequisites](#prerequisites)
  - [Setup Steps](#setup-steps)
  - [Running for Development](#running-for-development)
  - [Using ngrok for Webhook Testing](#using-ngrok-for-webhook-testing)
  - [Configuring Microsoft Entra (Azure AD) for Local Development](#configuring-microsoft-entra-azure-ad-for-local-development)
- [Support](#support)
- [Contributing](#contributing)
- [Code of Conduct](#code-of-conduct)
- [About Checkfirst](#about-checkfirst)
- [License](#license)

<!-- END doctoc generated TOC please keep comment here to allow auto update -->

## Features

- 🔗 Simple Microsoft authentication integration
- 📅 Calendar management (create/update/delete events)
- 📧 Email sending with rich content
- 🔔 Real-time notifications via webhooks
- 🔍 Event-driven architecture for easy integration

## Installation

```bash
npm install @checkfirst/nestjs-outlook
```

## Setup

### 1. Database Setup

This library requires database tables to store authentication and subscription data. You can use the built-in migrations to set up these tables automatically.

For details, see the [Migration Guide](src/migrations/README.md).

Alternatively, you can create the tables manually based on your database dialect (PostgreSQL, MySQL, etc.):

```typescript
import { MigrationInterface, QueryRunner } from 'typeorm';

export class CreateOutlookTables1697025846000 implements MigrationInterface {
  public async up(queryRunner: QueryRunner): Promise<void> {
    // Create required tables for webhooks, authentication, and user data
    // See the Migration Guide for details
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    // Drop tables in reverse order
    await queryRunner.query(`DROP TABLE IF EXISTS outlook_webhook_subscriptions`);
    await queryRunner.query(`DROP TABLE IF EXISTS microsoft_csrf_tokens`);
    await queryRunner.query(`DROP TABLE IF EXISTS microsoft_users`);
  }
}
```

### 2. Microsoft App Registration

Register your application with Microsoft to get the necessary credentials:

1. Go to the [Azure Portal](https://portal.azure.com/)
2. Navigate to Azure Active Directory > App registrations
3. Create a new registration
4. Configure redirects to include your callback URL (e.g., `https://your-api.example.com/auth/microsoft/callback`)
5. Add Microsoft Graph API permissions based on what features you need:
   - `Calendars.ReadWrite` - For calendar features
   - `Mail.Send` - For email features
   - `offline_access` - Required for all applications

### 3. Import Required Modules

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

### 4. Create an Auth Controller

The library provides a MicrosoftAuthController for handling authentication, but you can also create your own:

```typescript
import { Controller, Get, Req } from '@nestjs/common';
import { MicrosoftAuthService, PermissionScope } from '@checkfirst/nestjs-outlook';

@Controller('auth')
export class AuthController {
  constructor(
    private readonly microsoftAuthService: MicrosoftAuthService
  ) {}

  @Get('microsoft/login')
  async login(@Req() req: any) {
    // In a real application, get the user ID from your authentication system
    const userId = req.user?.id.toString() || '1';
        
    // Get the login URL with specific permission scopes
    return await this.microsoftAuthService.getLoginUrl(userId, [
      PermissionScope.CALENDAR_READ,
      PermissionScope.EMAIL_READ,
      PermissionScope.EMAIL_SEND
    ]);
  }
}
```

## Permission Scopes

You can request specific Microsoft permissions based on what your application needs:

```typescript
import { PermissionScope } from '@checkfirst/nestjs-outlook';

// Available permission scopes:
PermissionScope.CALENDAR_READ      // Read-only access to calendars
PermissionScope.CALENDAR_WRITE     // Read-write access to calendars
PermissionScope.EMAIL_READ         // Read-only access to emails
PermissionScope.EMAIL_WRITE        // Read-write access to emails
PermissionScope.EMAIL_SEND         // Permission to send emails
```

When getting a login URL, specify which permissions you need:

```typescript
// For a calendar-only app
const loginUrl = await microsoftAuthService.getLoginUrl(userId, [
  PermissionScope.CALENDAR_READ
]);

// For an email-only app
const loginUrl = await microsoftAuthService.getLoginUrl(userId, [
  PermissionScope.EMAIL_SEND
]);
```

## Available Services and Controllers

The library provides specialized services and controllers for Microsoft integration:

### 1. MicrosoftAuthService and MicrosoftAuthController

Handles the authentication flow with Microsoft:

```typescript
// Get a login URL to redirect your user to Microsoft's OAuth page
const loginUrl = await microsoftAuthService.getLoginUrl(userId);
```

After the user authenticates with Microsoft, they'll be redirected to your callback URL where you can complete the process.

### 2. CalendarService and CalendarController

Manage calendar operations:

```typescript
// Create a calendar event
const event = {
  subject: 'Team Meeting',
  start: {
    dateTime: '2023-06-01T10:00:00',
    timeZone: 'UTC',
  },
  end: {
    dateTime: '2023-06-01T11:00:00',
    timeZone: 'UTC',
  },
};

// Create the event
const result = await calendarService.createEvent(
  event,
  externalUserId,
  calendarId
);

// Get user's default calendar ID
const calendarId = await calendarService.getDefaultCalendarId(externalUserId);
```

The CalendarController provides a webhook endpoint at `/calendar/webhook` for receiving notifications from Microsoft Graph about calendar changes.

### 3. EmailService

Send emails:

```typescript
// Create email message
const message = {
  subject: 'Hello from NestJS Outlook',
  body: {
    contentType: 'HTML',
    content: '<p>This is the email body</p>'
  },
  toRecipients: [
    {
      emailAddress: {
        address: 'recipient@example.com'
      }
    }
  ]
};

// Send the email
const result = await emailService.sendEmail(
  message,
  externalUserId
);
```

## Events

The library emits events for various Microsoft activities that you can listen to in your application.

### Available Events

- `USER_AUTHENTICATED` - When a user completes authentication with Microsoft
- `EVENT_CREATED` - When a new calendar event is created 
- `EVENT_UPDATED` - When a calendar event is updated
- `EVENT_DELETED` - When a calendar event is deleted
- `EMAIL_RECEIVED` - When a new email is received
- `EMAIL_UPDATED` - When an email is updated
- `EMAIL_DELETED` - When an email is deleted

### Listening to Events

```typescript
import { Injectable } from '@nestjs/common';
import { OnEvent } from '@nestjs/event-emitter';
import { OutlookEventTypes, OutlookResourceData } from '@checkfirst/nestjs-outlook';

@Injectable()
export class YourService {
  // Handle user authentication event
  @OnEvent(OutlookEventTypes.USER_AUTHENTICATED)
  async handleUserAuthenticated(externalUserId: string, data: { externalUserId: string, scopes: string[] }) {
    console.log(`User ${externalUserId} authenticated with Microsoft`);
    // Perform any custom logic needed when a user authenticates
  }

  // Handle calendar events
  @OnEvent(OutlookEventTypes.EVENT_CREATED)
  handleOutlookEventCreated(data: OutlookResourceData) {
    console.log('New calendar event created:', data.id);
    // Handle the new event
  }
}
```

## Example Application Architecture

For a more complete example of how to structure your application using this library, check out the sample application included in this repository:

**Path:** `samples/nestjs-outlook-example/`

The sample app demonstrates a modular architecture with clear separation of concerns:

```
src/
├── auth/
│   ├── auth.controller.ts   # Handles Microsoft login and OAuth callback
│   └── auth.module.ts       # Configures MicrosoftOutlookModule for auth
│
├── calendar/
│   ├── calendar.controller.ts # API endpoints for calendar operations
│   ├── calendar.module.ts     # Configures MicrosoftOutlookModule for calendar
│   ├── calendar.service.ts    # Your business logic for calendars
│   └── dto/
│       └── create-event.dto.ts # Data validation for event creation
│
├── email/
│   ├── email.controller.ts   # API endpoints for email operations
│   ├── email.module.ts       # Configures MicrosoftOutlookModule for email
│   ├── email.service.ts      # Your business logic for emails
│   └── dto/
│       └── send-email.dto.ts # Data validation for email sending
│
└── app.module.ts             # Root module that imports feature modules
```

> **See:** [`samples/nestjs-outlook-example`](./samples/nestjs-outlook-example) for a full working example.

This modular architecture keeps concerns separated and makes your application easier to maintain and test.

## Local Development

This section provides instructions for developers who want to run and test the library locally with the sample application.

### Prerequisites

1. Install [yalc](https://github.com/wclr/yalc) globally:
   ```bash
   npm install -g yalc
   ```

2. Node.js and npm installed on your system

### Setup Steps

1. **Clone the repository:**
   ```bash
   git clone https://github.com/checkfirst-ltd/nestjs-outlook.git
   cd nestjs-outlook
   ```

2. **Install dependencies in the main library:**
   ```bash
   npm install
   ```

3. **Install dependencies in the sample application:**
   ```bash
   cd samples/nestjs-outlook-example
   npm install
   ```

4. **Link the library with the sample app using yalc:**
   
   In the root directory of the library:
   ```bash
   npm run build
   yalc publish
   ```
   
   In the sample app directory:
   ```bash
   yalc add @checkfirst/nestjs-outlook
   ```

### Running for Development

1. **Start the library in development mode:**
   
   In the root directory:
   ```bash
   npm run dev
   ```
   
   This will watch for changes in the library code and rebuild automatically.

2. **Start the sample app in development mode:**
   
   In the sample app directory:
   ```bash
   npm run start:dev:yalc
   ```
   
   This will start the NestJS sample application with hot-reload enabled, using the locally linked library.

### Using ngrok for Webhook Testing

To test Microsoft webhooks locally, you'll need to expose your local server to the internet using ngrok:

1. **Install ngrok** if you haven't already (https://ngrok.com/download)

2. **Start ngrok** to create a tunnel to your local server:
   ```bash
   ngrok http 3000
   ```

3. **Update your application configuration** with the ngrok URL:
   - Copy the HTTPS URL provided by ngrok (e.g., `https://1234-abcd-5678.ngrok.io`)
   - Update the `backendBaseUrl` in your `MicrosoftOutlookModule.forRoot()` configuration

4. **Register your webhook** with the Microsoft Graph API using the ngrok URL

This allows Microsoft to send webhook notifications to your local development environment.

### Configuring Microsoft Entra (Azure AD) for Local Development

For authentication to work properly with Microsoft, you need to configure your redirect URLs in Microsoft Entra (formerly Azure AD):

1. **Log in to the [Microsoft Entra admin center](https://entra.microsoft.com)** (or [Azure Portal](https://portal.azure.com) > Azure Active Directory)

2. **Navigate to App Registrations** and select your application

3. **Go to Authentication > Add a platform > Web**

4. **Add redirect URIs** for local development:
   - For local development with ngrok: `https://[your-ngrok-url]/auth/microsoft/callback`
   - For local development without ngrok: `http://localhost:[your-port]/auth/microsoft/callback`

5. **Save your changes**

> **Note:** The sample app runs on the port specified in your configuration (typically 3000). If you're using a different port, make sure to:
> - Start ngrok with your actual port: `ngrok http [your-port]`
> - Update your redirect URLs in Microsoft Entra accordingly
> - Update the port in your `.env` file or configuration

## Support

- [GitHub Issues](https://github.com/checkfirst-ltd/nestjs-outlook/issues)
- [Documentation](https://github.com/checkfirst-ltd/nestjs-outlook#readme)

## Contributing

We welcome contributions! Please see our [Contributing Guide](CONTRIBUTING.md) for more details.

## Code of Conduct

This project adheres to a Code of Conduct that all participants are expected to follow. Please read the [Code of Conduct](https://github.com/checkfirst-ltd/nestjs-outlook/blob/main/CONTRIBUTING.md#code-of-conduct) for details on our expectations.

## About Checkfirst

<a href="https://checkfirst.ai" target="_blank">
    <img src="https://raw.githubusercontent.com/checkfirst-ltd/nestjs-outlook/main/assets/checkfirst-logo.png" width="400" alt="CheckFirst Logo" />
</a>

[Checkfirst](https://checkfirst.ai) is a trusted provider of developer tools and solutions. We build open-source libraries that help developers create better applications faster.

## License

[MIT](LICENSE) © [CheckFirst Ltd](https://checkfirst.ai)
