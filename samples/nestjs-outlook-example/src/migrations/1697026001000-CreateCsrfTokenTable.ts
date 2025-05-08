import { MigrationInterface, QueryRunner } from 'typeorm';

export class CreateCsrfTokenTable1697026001000 implements MigrationInterface {
  public async up(queryRunner: QueryRunner): Promise<void> {
    // Check if the table exists
    const tableExists = await queryRunner.query(`
      SELECT COUNT(*) as count
      FROM sqlite_master
      WHERE type='table' 
      AND name='microsoft_csrf_tokens'
    `);

    if (tableExists[0].count === 0) {
      // Create the table if it doesn't exist
      await queryRunner.query(`
        CREATE TABLE microsoft_csrf_tokens (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          token VARCHAR(64) NOT NULL,
          user_id VARCHAR(255) NOT NULL,
          expires TIMESTAMP NOT NULL,
          created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL,
          CONSTRAINT "UQ_microsoft_csrf_tokens_token" UNIQUE (token)
        )
      `);
    }
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.query(`DROP TABLE IF EXISTS microsoft_csrf_tokens`);
  }
} 