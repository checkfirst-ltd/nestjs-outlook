import {
  MigrationInterface,
  QueryRunner,
  Table,
  TableIndex,
  TableForeignKey,
} from 'typeorm';

export class AddTenantUsersTable1782207500000 implements MigrationInterface {
  public async up(queryRunner: QueryRunner): Promise<void> {
    // Create tenant_users table
    await queryRunner.createTable(
      new Table({
        name: 'tenant_users',
        columns: [
          {
            name: 'id',
            type: 'INTEGER',
            isPrimary: true,
            isGenerated: true,
            generationStrategy: 'increment',
          },
          {
            name: 'tenant_connection_id',
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
            length: '320',
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

    // Create foreign key to microsoft_tenants (tenant_connection is an alias)
    await queryRunner.createForeignKey(
      'tenant_users',
      new TableForeignKey({
        name: 'FK_tenant_users_tenant_connection',
        columnNames: ['tenant_connection_id'],
        referencedTableName: 'microsoft_tenants',
        referencedColumnNames: ['id'],
        onDelete: 'CASCADE',
      })
    );

    // Create index on microsoft_user_id
    await queryRunner.createIndex(
      'tenant_users',
      new TableIndex({
        name: 'IDX_tenant_users_microsoft_user_id',
        columnNames: ['microsoft_user_id'],
      })
    );

    // Create index on external_user_id
    await queryRunner.createIndex(
      'tenant_users',
      new TableIndex({
        name: 'IDX_tenant_users_external_user_id',
        columnNames: ['external_user_id'],
      })
    );
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    // Drop indexes
    await queryRunner.dropIndex('tenant_users', 'IDX_tenant_users_external_user_id');
    await queryRunner.dropIndex('tenant_users', 'IDX_tenant_users_microsoft_user_id');

    // Drop foreign key
    const table = await queryRunner.getTable('tenant_users');
    if (table) {
      const foreignKey = table.foreignKeys.find(
        (fk) => fk.name === 'FK_tenant_users_tenant_connection'
      );
      if (foreignKey) {
        await queryRunner.dropForeignKey('tenant_users', foreignKey);
      }
    }

    // Drop table
    await queryRunner.dropTable('tenant_users');
  }
}
