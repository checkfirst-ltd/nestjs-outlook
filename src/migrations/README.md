# Microsoft Outlook Module Migrations

This directory contains the database migrations required for the NestJS Outlook module.

## Available Migrations

1. `1697025846000-CreateOutlookTables.ts` - Creates the initial database tables for the module
2. `1697026000000-EnsureUniqueSubscriptions.ts` - Ensures unique constraints on subscription IDs
3. `1699000000000-AddMicrosoftUserTable.ts` - Adds the Microsoft User table for improved token management

## Migration Order

Please run these migrations in order, as each migration builds upon the previous one.

## Migration Overview

### 1. CreateOutlookTables
- Creates the `outlook_webhook_subscriptions` table for storing webhook subscriptions
- Creates the `microsoft_csrf_tokens` table for CSRF protection during OAuth flow

### 2. EnsureUniqueSubscriptions
- Adds unique constraints to the `subscription_id` column in the `outlook_webhook_subscriptions` table

### 3. AddMicrosoftUserTable
- Adds a new `microsoft_users` table to properly store user tokens and scopes
- Removes token columns from the `outlook_webhook_subscriptions` table since they're now stored in the users table
- Centralizes token management for better security and easier refresh

## Running Migrations

In your NestJS application, ensure TypeORM is configured to run migrations. Example:

```typescript
// in your app.module.ts
TypeOrmModule.forRoot({
  // Your database configuration
  migrations: [__dirname + '/../migrations/*{.ts,.js}'],
  migrationsRun: true,
})
```

Or run migrations manually using the TypeORM CLI:

```bash
npx typeorm migration:run
``` 