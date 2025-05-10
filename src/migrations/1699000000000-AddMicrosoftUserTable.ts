import { MigrationInterface, QueryRunner, Table, TableIndex, TableColumn } from 'typeorm';

export class AddMicrosoftUserTable1699000000000 implements MigrationInterface {
  public async up(queryRunner: QueryRunner): Promise<void> {
    // Create microsoft_users table
    await queryRunner.createTable(
      new Table({
        name: 'microsoft_users',
        columns: [
          {
            name: 'id',
            type: 'INTEGER',
            isPrimary: true,
            isGenerated: true,
            generationStrategy: 'increment',
          },
          {
            name: 'external_user_id',
            type: 'varchar',
            length: '255',
            isNullable: false,
          },
          {
            name: 'access_token',
            type: 'text',
            isNullable: false,
          },
          {
            name: 'refresh_token',
            type: 'text',
            isNullable: false,
          },
          {
            name: 'token_expiry',
            type: 'timestamp',
            isNullable: false,
          },
          {
            name: 'scopes',
            type: 'text',
            isNullable: false,
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

    // Create index on external_user_id
    await queryRunner.createIndex(
      'microsoft_users',
      new TableIndex({
        name: 'IDX_microsoft_users_external_user_id',
        columnNames: ['external_user_id'],
      })
    );

    // Check if access_token and refresh_token columns exist in outlook_webhook_subscriptions
    const table = await queryRunner.getTable('outlook_webhook_subscriptions');
    
    if (table) {
      const accessTokenColumn = table.findColumnByName('access_token');
      const refreshTokenColumn = table.findColumnByName('refresh_token');

      // Drop columns if they exist
      if (accessTokenColumn) {
        await queryRunner.dropColumn('outlook_webhook_subscriptions', 'access_token');
      }
      
      if (refreshTokenColumn) {
        await queryRunner.dropColumn('outlook_webhook_subscriptions', 'refresh_token');
      }
    }
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    // Add back the token columns to outlook_webhook_subscriptions if they don't exist
    const table = await queryRunner.getTable('outlook_webhook_subscriptions');
    
    if (table) {
      const accessTokenColumn = table.findColumnByName('access_token');
      const refreshTokenColumn = table.findColumnByName('refresh_token');

      // Add columns if they don't exist
      if (!accessTokenColumn) {
        await queryRunner.addColumn(
          'outlook_webhook_subscriptions',
          new TableColumn({
            name: 'access_token',
            type: 'text',
            isNullable: true,
          })
        );
      }
      
      if (!refreshTokenColumn) {
        await queryRunner.addColumn(
          'outlook_webhook_subscriptions',
          new TableColumn({
            name: 'refresh_token',
            type: 'text',
            isNullable: true,
          })
        );
      }
    }

    // Drop the index and table
    await queryRunner.dropIndex('microsoft_users', 'IDX_microsoft_users_external_user_id');
    await queryRunner.dropTable('microsoft_users');
  }
} 