import { MigrationInterface, QueryRunner } from 'typeorm';

interface QueryCount {
  count: number;
}

export class EnsureUniqueSubscriptions1697026000000 implements MigrationInterface {
  public async up(queryRunner: QueryRunner): Promise<void> {
    // Check if the table exists
    const tableExists = (await queryRunner.query(`
      SELECT COUNT(*) as count
      FROM sqlite_master
      WHERE type='table' 
      AND name='outlook_webhook_subscriptions'
    `)) as QueryCount[];

    if (tableExists[0].count === 0) {
      // Create the table if it doesn't exist
      await queryRunner.query(`
        CREATE TABLE outlook_webhook_subscriptions (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          subscription_id VARCHAR(255) NOT NULL UNIQUE,
          user_id INTEGER NOT NULL,
          resource VARCHAR(255) NOT NULL,
          change_type VARCHAR(255) NOT NULL,
          client_state VARCHAR(255) NOT NULL,
          notification_url VARCHAR(255) NOT NULL,
          expiration_date_time TIMESTAMP NOT NULL,
          is_active BOOLEAN DEFAULT true,
          access_token TEXT,
          refresh_token TEXT,
          created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL,
          updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL
        )
      `);
    } else {
      // Check if the unique constraint exists
      const uniqueConstraintExists = (await queryRunner.query(`
        SELECT COUNT(*) as count
        FROM sqlite_master
        WHERE type='index' 
        AND tbl_name='outlook_webhook_subscriptions' 
        AND name='UQ_outlook_webhook_subscriptions_id'
      `)) as QueryCount[];

      // If the constraint doesn't exist, add it
      if (uniqueConstraintExists[0].count === 0) {
        await queryRunner.query(`
          CREATE UNIQUE INDEX IF NOT EXISTS "UQ_outlook_webhook_subscriptions_id" 
          ON "outlook_webhook_subscriptions" ("subscription_id")
        `);
      }

      // Handle duplicate subscription IDs by keeping only the most recent one
      await queryRunner.query(`
        -- Create a temporary table with unique subscription_ids
        CREATE TEMPORARY TABLE IF NOT EXISTS tmp_outlook_subscriptions AS
        SELECT * FROM (
          SELECT * FROM outlook_webhook_subscriptions
          ORDER BY created_at DESC, id DESC
        ) GROUP BY subscription_id;

        -- Delete all records from the original table
        DELETE FROM outlook_webhook_subscriptions;

        -- Reset the autoincrement counter
        DELETE FROM sqlite_sequence WHERE name='outlook_webhook_subscriptions';

        -- Reinsert the unique records
        INSERT INTO outlook_webhook_subscriptions
        SELECT * FROM tmp_outlook_subscriptions;

        -- Drop the temporary table
        DROP TABLE tmp_outlook_subscriptions;
      `);
    }
  }

  public async down(_queryRunner: QueryRunner): Promise<void> {
    // This migration should not be reversed, as it fixes a data integrity issue
  }
} 