import { MigrationInterface, QueryRunner } from 'typeorm';

interface QueryResult {
  count: number;
}

export class EnsureUniqueSubscriptions1697026000000 implements MigrationInterface {
  public async up(queryRunner: QueryRunner): Promise<void> {
    // Check if the unique constraint exists
    const uniqueConstraintExists = await queryRunner.query(`
      SELECT COUNT(*) as count
      FROM sqlite_master
      WHERE type='table' 
      AND name='outlook_webhook_subscriptions' 
      AND sql LIKE '%CONSTRAINT "UQ_outlook_webhook_subscriptions_id"%'
    `) as QueryResult[];

    // If the constraint doesn't exist, add it
    if (uniqueConstraintExists[0].count === 0) {
      // First, add the constraint if it doesn't exist
      await queryRunner.query(`
        CREATE UNIQUE INDEX IF NOT EXISTS "UQ_outlook_webhook_subscriptions_id" 
        ON "outlook_webhook_subscriptions" ("subscription_id")
      `);
    }
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    // Drop the unique index if it exists
    await queryRunner.query(`
      DROP INDEX IF EXISTS "UQ_outlook_webhook_subscriptions_id"
    `);
  }
} 