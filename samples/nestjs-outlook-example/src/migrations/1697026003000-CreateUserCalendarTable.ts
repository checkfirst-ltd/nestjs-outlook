import { MigrationInterface, QueryRunner } from 'typeorm';

interface QueryCount {
  count: number;
}

export class CreateUserCalendarTable1697026003000 implements MigrationInterface {
  public async up(queryRunner: QueryRunner): Promise<void> {
    // Check if the table exists
    const tableExists = (await queryRunner.query(`
      SELECT COUNT(*) as count
      FROM sqlite_master
      WHERE type='table' 
      AND name='user_calendars'
    `)) as QueryCount[];

    if (tableExists[0].count === 0) {
      // Create the user_calendars table
      await queryRunner.query(`
        CREATE TABLE user_calendars (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          user_id INTEGER NOT NULL,
          calendar_id VARCHAR(255) NOT NULL,
          access_token TEXT NOT NULL,
          refresh_token TEXT NOT NULL,
          token_expiry TIMESTAMP NOT NULL,
          active BOOLEAN DEFAULT true,
          created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL,
          updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL,
          FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
        )
      `);

      // Create compound index on user_id and calendar_id
      await queryRunner.query(`
        CREATE UNIQUE INDEX IF NOT EXISTS "IDX_user_calendars_user_id_calendar_id" 
        ON "user_calendars" ("user_id", "calendar_id")
      `);
    }
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.query(`DROP TABLE IF EXISTS user_calendars`);
  }
} 