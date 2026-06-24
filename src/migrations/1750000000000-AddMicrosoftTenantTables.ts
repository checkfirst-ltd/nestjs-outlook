import {
  MigrationInterface,
  QueryRunner,
  Table,
  TableIndex,
  TableColumn,
  TableForeignKey,
} from 'typeorm';

export class AddMicrosoftTenantTables1750000000000 implements MigrationInterface {
  public async up(queryRunner: QueryRunner): Promise<void> {
    // Create microsoft_tenants table
    await queryRunner.createTable(
      new Table({
        name: 'microsoft_tenants',
        columns: [
          {
            name: 'id',
            type: 'INTEGER',
            isPrimary: true,
            isGenerated: true,
            generationStrategy: 'increment',
          },
          {
            name: 'tenant_id',
            type: 'varchar',
            length: '36',
            isNullable: false,
            isUnique: true,
          },
          {
            name: 'client_id',
            type: 'varchar',
            length: '36',
            isNullable: false,
          },
          {
            name: 'certificate_thumbprint',
            type: 'varchar',
            length: '64',
            isNullable: false,
          },
          {
            name: 'certificate_path',
            type: 'varchar',
            length: '255',
            isNullable: true,
          },
          {
            name: 'certificate_key_path',
            type: 'varchar',
            length: '255',
            isNullable: true,
          },
          {
            name: 'status',
            type: 'varchar',
            length: '32',
            default: "'PENDING_CONSENT'",
            isNullable: false,
          },
          {
            name: 'admin_consent_granted_at',
            type: 'datetime',
            isNullable: true,
          },
          {
            name: 'is_active',
            type: 'boolean',
            default: true,
            isNullable: false,
          },
          {
            name: 'created_at',
            type: 'timestamp',
            default: 'now()',
            isNullable: false,
          },
          {
            name: 'updated_at',
            type: 'timestamp',
            default: 'now()',
            isNullable: false,
          },
        ],
      }),
      true
    );

    // Create index on tenant_id (unique already enforced by column)
    await queryRunner.createIndex(
      'microsoft_tenants',
      new TableIndex({
        name: 'IDX_microsoft_tenants_tenant_id',
        columnNames: ['tenant_id'],
      })
    );

    // Create microsoft_tenant_users table
    await queryRunner.createTable(
      new Table({
        name: 'microsoft_tenant_users',
        columns: [
          {
            name: 'id',
            type: 'INTEGER',
            isPrimary: true,
            isGenerated: true,
            generationStrategy: 'increment',
          },
          {
            name: 'tenant_id',
            type: 'INTEGER',
            isNullable: false,
          },
          {
            name: 'microsoft_user_id',
            type: 'varchar',
            length: '36',
            isNullable: false,
          },
          {
            name: 'external_user_id',
            type: 'varchar',
            length: '255',
            isNullable: false,
          },
          {
            name: 'user_principal_name',
            type: 'varchar',
            length: '255',
            isNullable: false,
          },
          {
            name: 'default_calendar_id',
            type: 'varchar',
            length: '255',
            isNullable: true,
          },
          {
            name: 'is_active',
            type: 'boolean',
            default: true,
            isNullable: false,
          },
          {
            name: 'created_at',
            type: 'timestamp',
            default: 'now()',
            isNullable: false,
          },
          {
            name: 'updated_at',
            type: 'timestamp',
            default: 'now()',
            isNullable: false,
          },
        ],
      }),
      true
    );

    // Create foreign key from microsoft_tenant_users to microsoft_tenants
    await queryRunner.createForeignKey(
      'microsoft_tenant_users',
      new TableForeignKey({
        name: 'FK_microsoft_tenant_users_tenant',
        columnNames: ['tenant_id'],
        referencedTableName: 'microsoft_tenants',
        referencedColumnNames: ['id'],
        onDelete: 'CASCADE',
      })
    );

    // Create index on external_user_id
    await queryRunner.createIndex(
      'microsoft_tenant_users',
      new TableIndex({
        name: 'IDX_microsoft_tenant_users_external_user_id',
        columnNames: ['external_user_id'],
      })
    );

    // Create index on microsoft_user_id
    await queryRunner.createIndex(
      'microsoft_tenant_users',
      new TableIndex({
        name: 'IDX_microsoft_tenant_users_microsoft_user_id',
        columnNames: ['microsoft_user_id'],
      })
    );

    // Add tenant_id column to outlook_webhook_subscriptions
    await queryRunner.addColumn(
      'outlook_webhook_subscriptions',
      new TableColumn({
        name: 'tenant_id',
        type: 'INTEGER',
        isNullable: true,
      })
    );

    // Add microsoft_user_id column to outlook_webhook_subscriptions
    await queryRunner.addColumn(
      'outlook_webhook_subscriptions',
      new TableColumn({
        name: 'microsoft_user_id',
        type: 'varchar',
        length: '36',
        isNullable: true,
      })
    );

    // Create foreign key from outlook_webhook_subscriptions to microsoft_tenants
    await queryRunner.createForeignKey(
      'outlook_webhook_subscriptions',
      new TableForeignKey({
        name: 'FK_outlook_webhook_subscriptions_tenant',
        columnNames: ['tenant_id'],
        referencedTableName: 'microsoft_tenants',
        referencedColumnNames: ['id'],
        onDelete: 'SET NULL',
      })
    );
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    // Drop foreign key from outlook_webhook_subscriptions
    const subscriptionsTable = await queryRunner.getTable('outlook_webhook_subscriptions');
    if (subscriptionsTable) {
      const foreignKey = subscriptionsTable.foreignKeys.find(
        (fk) => fk.name === 'FK_outlook_webhook_subscriptions_tenant'
      );
      if (foreignKey) {
        await queryRunner.dropForeignKey('outlook_webhook_subscriptions', foreignKey);
      }
    }

    // Drop columns from outlook_webhook_subscriptions
    await queryRunner.dropColumn('outlook_webhook_subscriptions', 'microsoft_user_id');
    await queryRunner.dropColumn('outlook_webhook_subscriptions', 'tenant_id');

    // Drop indexes from microsoft_tenant_users
    await queryRunner.dropIndex(
      'microsoft_tenant_users',
      'IDX_microsoft_tenant_users_microsoft_user_id'
    );
    await queryRunner.dropIndex(
      'microsoft_tenant_users',
      'IDX_microsoft_tenant_users_external_user_id'
    );

    // Drop foreign key from microsoft_tenant_users
    const tenantUsersTable = await queryRunner.getTable('microsoft_tenant_users');
    if (tenantUsersTable) {
      const foreignKey = tenantUsersTable.foreignKeys.find(
        (fk) => fk.name === 'FK_microsoft_tenant_users_tenant'
      );
      if (foreignKey) {
        await queryRunner.dropForeignKey('microsoft_tenant_users', foreignKey);
      }
    }

    // Drop microsoft_tenant_users table
    await queryRunner.dropTable('microsoft_tenant_users');

    // Drop index from microsoft_tenants
    await queryRunner.dropIndex('microsoft_tenants', 'IDX_microsoft_tenants_tenant_id');

    // Drop microsoft_tenants table
    await queryRunner.dropTable('microsoft_tenants');
  }
}
