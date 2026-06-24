import { MigrationInterface, QueryRunner } from 'typeorm';

export class AddMicrosoftUserStatusColumn1699000000002 implements MigrationInterface {
  public async up(queryRunner: QueryRunner): Promise<void> {
    // Check if the table exists
    const tableExists = await queryRunner.query(`
      SELECT COUNT(*) as count
      FROM sqlite_master
      WHERE type='table'
      AND name='microsoft_users'
    `);

    if (tableExists[0].count > 0) {
      // Check if status column exists
      const columns = await queryRunner.query(`PRAGMA table_info(microsoft_users)`);
      const hasStatus = columns.some((col: { name: string }) => col.name === 'status');
      const hasDefaultCalendarId = columns.some((col: { name: string }) => col.name === 'default_calendar_id');

      if (!hasStatus) {
        await queryRunner.query(`
          ALTER TABLE microsoft_users ADD COLUMN status VARCHAR(32) DEFAULT 'ACTIVE' NOT NULL
        `);
      }

      if (!hasDefaultCalendarId) {
        await queryRunner.query(`
          ALTER TABLE microsoft_users ADD COLUMN default_calendar_id VARCHAR(255) NULL
        `);
      }
    }
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    // SQLite doesn't support DROP COLUMN directly, so we skip this
  }
}
