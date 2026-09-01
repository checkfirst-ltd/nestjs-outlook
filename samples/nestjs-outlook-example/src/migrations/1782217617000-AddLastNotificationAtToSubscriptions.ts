import { MigrationInterface, QueryRunner } from 'typeorm';

export class AddLastNotificationAtToSubscriptions1782217617000 implements MigrationInterface {
  public async up(queryRunner: QueryRunner): Promise<void> {
    // Add last_notification_at column to outlook_webhook_subscriptions.
    // Mirrors the package migration AddLastNotificationAtToSubscriptions1776600000000,
    // which the sample app's local migration set was missing.
    const subscriptionsTable = await queryRunner.getTable('outlook_webhook_subscriptions');
    if (subscriptionsTable && !subscriptionsTable.findColumnByName('last_notification_at')) {
      await queryRunner.query(`
        ALTER TABLE outlook_webhook_subscriptions
        ADD COLUMN last_notification_at DATETIME
      `);
    }
  }

  public async down(): Promise<void> {
    // SQLite doesn't support DROP COLUMN directly. The column is nullable and unused
    // when reverted, so we leave it in place.
  }
}
