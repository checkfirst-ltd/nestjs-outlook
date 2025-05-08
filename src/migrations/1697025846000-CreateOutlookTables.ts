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
        updated_at TIMESTAMP DEFAULT NOW() NOT NULL,
        CONSTRAINT "UQ_outlook_webhook_subscriptions_id" UNIQUE (subscription_id)
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