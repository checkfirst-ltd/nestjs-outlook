import { MigrationInterface, QueryRunner } from 'typeorm';

interface QueryCount {
  count: number;
}

export class AddMicrosoftTenantTables1782217616000 implements MigrationInterface {
  public async up(queryRunner: QueryRunner): Promise<void> {
    // Create microsoft_tenants table
    const tenantsTableExists = (await queryRunner.query(`
      SELECT COUNT(*) as count
      FROM sqlite_master
      WHERE type='table'
      AND name='microsoft_tenants'
    `)) as QueryCount[];

    if (tenantsTableExists[0].count === 0) {
      await queryRunner.query(`
        CREATE TABLE microsoft_tenants (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          tenant_id VARCHAR(36) NOT NULL UNIQUE,
          client_id VARCHAR(36) NOT NULL,
          certificate_thumbprint VARCHAR(64) NOT NULL,
          certificate_path VARCHAR(255),
          certificate_key_path VARCHAR(255),
          status VARCHAR(32) NOT NULL DEFAULT 'PENDING_CONSENT',
          admin_consent_granted_at DATETIME,
          is_active BOOLEAN NOT NULL DEFAULT 1,
          created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL,
          updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL
        )
      `);

      // Create index on tenant_id
      await queryRunner.query(`
        CREATE INDEX IF NOT EXISTS "IDX_microsoft_tenants_tenant_id"
        ON "microsoft_tenants" ("tenant_id")
      `);
    }

    // Create microsoft_tenant_users table
    const tenantUsersTableExists = (await queryRunner.query(`
      SELECT COUNT(*) as count
      FROM sqlite_master
      WHERE type='table'
      AND name='microsoft_tenant_users'
    `)) as QueryCount[];

    if (tenantUsersTableExists[0].count === 0) {
      await queryRunner.query(`
        CREATE TABLE microsoft_tenant_users (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          tenant_id INTEGER NOT NULL,
          microsoft_user_id VARCHAR(36) NOT NULL,
          external_user_id VARCHAR(255) NOT NULL,
          user_principal_name VARCHAR(255) NOT NULL,
          default_calendar_id VARCHAR(255),
          is_active BOOLEAN NOT NULL DEFAULT 1,
          created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL,
          updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL,
          FOREIGN KEY (tenant_id) REFERENCES microsoft_tenants(id) ON DELETE CASCADE
        )
      `);

      // Create index on external_user_id
      await queryRunner.query(`
        CREATE INDEX IF NOT EXISTS "IDX_microsoft_tenant_users_external_user_id"
        ON "microsoft_tenant_users" ("external_user_id")
      `);

      // Create index on microsoft_user_id
      await queryRunner.query(`
        CREATE INDEX IF NOT EXISTS "IDX_microsoft_tenant_users_microsoft_user_id"
        ON "microsoft_tenant_users" ("microsoft_user_id")
      `);
    }

    // Add tenant_id and microsoft_user_id columns to outlook_webhook_subscriptions
    const subscriptionsTable = await queryRunner.getTable('outlook_webhook_subscriptions');
    if (subscriptionsTable) {
      const hasTenantId = subscriptionsTable.findColumnByName('tenant_id');
      const hasMicrosoftUserId = subscriptionsTable.findColumnByName('microsoft_user_id');

      if (!hasTenantId) {
        await queryRunner.query(`
          ALTER TABLE outlook_webhook_subscriptions
          ADD COLUMN tenant_id INTEGER REFERENCES microsoft_tenants(id) ON DELETE SET NULL
        `);
      }

      if (!hasMicrosoftUserId) {
        await queryRunner.query(`
          ALTER TABLE outlook_webhook_subscriptions
          ADD COLUMN microsoft_user_id VARCHAR(36)
        `);
      }
    }
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    // SQLite doesn't support DROP COLUMN directly, so we need to recreate the table
    // For simplicity, we'll just drop the new tables and leave the subscription columns
    // (they'll be ignored if not used)

    // Drop microsoft_tenant_users table
    await queryRunner.query(`DROP TABLE IF EXISTS microsoft_tenant_users`);

    // Drop microsoft_tenants table
    await queryRunner.query(`DROP TABLE IF EXISTS microsoft_tenants`);

    // Note: SQLite doesn't support DROP COLUMN, so tenant_id and microsoft_user_id
    // columns will remain in outlook_webhook_subscriptions. They will be null and unused.
  }
}
