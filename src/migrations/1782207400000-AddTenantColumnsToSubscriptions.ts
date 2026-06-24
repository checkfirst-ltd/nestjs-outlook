import { MigrationInterface, QueryRunner, TableColumn, TableIndex } from 'typeorm';

/**
 * Adds tenantId and microsoftUserId columns to outlook_webhook_subscriptions table.
 * These columns support app-only (tenant-wide) subscriptions that use
 * /users/{microsoftUserId}/events instead of /me/events.
 */
export class AddTenantColumnsToSubscriptions1782207400000 implements MigrationInterface {
  name = 'AddTenantColumnsToSubscriptions1782207400000';

  public async up(queryRunner: QueryRunner): Promise<void> {
    // Add tenant_id column
    await queryRunner.addColumn(
      'outlook_webhook_subscriptions',
      new TableColumn({
        name: 'tenant_id',
        type: 'varchar',
        length: '36',
        isNullable: true,
        comment: 'Microsoft tenant ID for app-only subscriptions',
      }),
    );

    // Add microsoft_user_id column
    await queryRunner.addColumn(
      'outlook_webhook_subscriptions',
      new TableColumn({
        name: 'microsoft_user_id',
        type: 'varchar',
        length: '255',
        isNullable: true,
        comment: 'Microsoft user ID (immutable ID) for app-only subscriptions',
      }),
    );

    // Add index on tenant_id for efficient tenant-based queries
    await queryRunner.createIndex(
      'outlook_webhook_subscriptions',
      new TableIndex({
        name: 'IDX_outlook_webhook_subscriptions_tenant_id',
        columnNames: ['tenant_id'],
      }),
    );

    // Add index on microsoft_user_id for efficient user-based queries
    await queryRunner.createIndex(
      'outlook_webhook_subscriptions',
      new TableIndex({
        name: 'IDX_outlook_webhook_subscriptions_microsoft_user_id',
        columnNames: ['microsoft_user_id'],
      }),
    );
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    // Drop indexes first
    await queryRunner.dropIndex(
      'outlook_webhook_subscriptions',
      'IDX_outlook_webhook_subscriptions_microsoft_user_id',
    );
    await queryRunner.dropIndex(
      'outlook_webhook_subscriptions',
      'IDX_outlook_webhook_subscriptions_tenant_id',
    );

    // Drop columns
    await queryRunner.dropColumn('outlook_webhook_subscriptions', 'microsoft_user_id');
    await queryRunner.dropColumn('outlook_webhook_subscriptions', 'tenant_id');
  }
}
