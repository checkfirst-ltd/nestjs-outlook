---
dep:
  type: tutorial
  audience: [library-integrator]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: medium
  depends_on:
    - src/microsoft-outlook.module.ts
    - src/interfaces/config/outlook-config.interface.ts
    - src/controllers/microsoft-auth.controller.ts
    - README.md
  tags: [getting-started, onboarding, oauth, setup]
  links:
    - target: ../reference/configuration.md
      rel: TEACHES
    - target: ../reference/permission-scopes.md
      rel: TEACHES
    - target: ../reference/http-endpoints.md
      rel: TEACHES
    - target: ../reference/microsoft-auth-service.md
      rel: TEACHES
    - target: ../how-to/authenticate-a-user.md
      rel: NEXT
    - target: ../how-to/index.md
      rel: NEXT
---

# Getting Started: Connect Your First User to Microsoft

By the end of this tutorial you will have a running NestJS application that can send a
user to Microsoft, receive them back after they sign in, and store their tokens so the
module can act on their behalf.

## Prerequisites

Before you begin, make sure you have:

- A NestJS application (v10) that already boots with `npm run start:dev`.
- A working TypeORM connection (`@nestjs/typeorm`) to a SQL database (PostgreSQL or MySQL).
- `@nestjs/schedule` and `@nestjs/event-emitter` installed (the module registers cron jobs and emits events).
- A Microsoft account with access to [Microsoft Entra](https://entra.microsoft.com/) so you can register an app.

## Step 1 — Install the package

```bash
npm install @checkfirst/nestjs-outlook
```

**Expected result:** `@checkfirst/nestjs-outlook` appears in your `package.json` dependencies.

## Step 2 — Create the database tables

The module persists Microsoft users, CSRF tokens, webhook subscriptions, and delta links.
Run the bundled migrations against your database.

Point your TypeORM migration configuration at the package's compiled migrations, then run:

```bash
npm run typeorm migration:run
```

**Expected result:** Your database now contains `microsoft_users`, `microsoft_csrf_tokens`,
`outlook_webhook_subscriptions`, and `outlook_delta_links` tables.

> The full migration setup is described in the package's `src/migrations/README.md`.

## Step 3 — Register your application with Microsoft

1. Open [Microsoft Entra](https://entra.microsoft.com/) and go to **App registrations**.
2. Create a new registration.
3. Under **Authentication → Add a platform → Web**, add a redirect URI that matches the
   callback your app will expose, for example `https://your-api.example.com/api/v1/auth/microsoft/callback`.
4. Under **API permissions**, add the Microsoft Graph delegated permissions you need
   (for example `Calendars.ReadWrite`, `Mail.Send`, and `offline_access`).
5. Under **Certificates & secrets**, create a client secret and copy its value.

**Expected result:** You have a **client ID**, a **client secret**, and a **redirect URI**
registered with Microsoft.

## Step 4 — Import the module

Register `MicrosoftOutlookModule` in your root module and supply the credentials from Step 3.

```typescript
import { Module } from '@nestjs/common';
import { TypeOrmModule } from '@nestjs/typeorm';
import { ScheduleModule } from '@nestjs/schedule';
import { EventEmitterModule } from '@nestjs/event-emitter';
import { MicrosoftOutlookModule } from '@checkfirst/nestjs-outlook';

@Module({
  imports: [
    ScheduleModule.forRoot(),
    EventEmitterModule.forRoot(),
    MicrosoftOutlookModule.forRoot({
      clientId: process.env.MS_CLIENT_ID,
      clientSecret: process.env.MS_CLIENT_SECRET,
      redirectPath: 'auth/microsoft/callback',
      backendBaseUrl: 'https://your-api.example.com',
      basePath: 'api/v1',
    }),
  ],
})
export class AppModule {}
```

Make sure the module's entities are included in your TypeORM `entities` glob so the tables
from Step 2 are mapped.

**Expected result:** The app starts and logs that the Outlook state backend is `in-memory`.

> Every configuration field, and the `forRootAsync` variant, is listed in the
> [Configuration reference](../reference/configuration.md).

## Step 5 — Add a login route

The module ships a `MicrosoftAuthService`. Inject it and build a login URL that redirects
the user to Microsoft.

```typescript
import { Controller, Get, Req, Res } from '@nestjs/common';
import { Response } from 'express';
import { MicrosoftAuthService, PermissionScope } from '@checkfirst/nestjs-outlook';

@Controller('auth')
export class AuthController {
  constructor(private readonly microsoftAuthService: MicrosoftAuthService) {}

  @Get('microsoft/login')
  async login(@Req() req: any, @Res() res: Response) {
    const userId = req.user?.id?.toString() ?? '1';
    const loginUrl = await this.microsoftAuthService.getLoginUrl(userId, [
      PermissionScope.CALENDAR_READ,
      PermissionScope.EMAIL_SEND,
    ]);
    return res.redirect(loginUrl);
  }
}
```

**Expected result:** Visiting `/auth/microsoft/login` redirects your browser to the
Microsoft sign-in page.

> The scope values come from the [Permission scopes reference](../reference/permission-scopes.md).

## Step 6 — Complete the sign-in

The module already exposes the callback controller at `auth/microsoft/callback` (under your
`basePath`). When the user finishes signing in, Microsoft redirects them there, the module
exchanges the authorization code for tokens, and stores them.

Sign in with your Microsoft account and approve the requested permissions.

**Expected result:** A row appears in `microsoft_users` for your account, and the module
emits a `USER_AUTHENTICATED` event.

> The callback route and its query parameters are documented in the
> [HTTP endpoints reference](../reference/http-endpoints.md).

## What you built

You now have a NestJS app that:

- Registers the Microsoft Outlook module with your Entra credentials.
- Redirects users into the Microsoft OAuth flow with the scopes you choose.
- Receives the callback, exchanges the code for tokens, and persists the connected user.

## Next steps

- [Authenticate a user](../how-to/authenticate-a-user.md) — the auth flow as a focused task.
- [Browse all how-to guides](../how-to/index.md) — send email, manage calendar events, and set up webhooks.
