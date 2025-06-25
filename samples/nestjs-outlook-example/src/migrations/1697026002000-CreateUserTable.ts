import { MigrationInterface, QueryRunner } from 'typeorm';

interface QueryCount {
  count: number;
}

export class CreateUserTable1697026002000 implements MigrationInterface {
  public async up(queryRunner: QueryRunner): Promise<void> {
    // Check if the table exists
    const tableExists = (await queryRunner.query(`
      SELECT COUNT(*) as count
      FROM sqlite_master
      WHERE type='table' 
      AND name='users'
    `)) as QueryCount[];

    if (tableExists[0].count === 0) {
      // Create the users table
      await queryRunner.query(`
        CREATE TABLE users (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          email VARCHAR(255) NOT NULL UNIQUE,
          name VARCHAR(255),
          created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL,
          updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL
        )
      `);

      // Create index on email field
      await queryRunner.query(`
        CREATE UNIQUE INDEX IF NOT EXISTS "IDX_users_email" 
        ON "users" ("email")
      `);

      // Insert a test user
      await queryRunner.query(`
        INSERT INTO users (email, name) 
        VALUES ('test@example.com', 'Test User')
      `);
    }
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.query(`DROP TABLE IF EXISTS users`);
  }
} 