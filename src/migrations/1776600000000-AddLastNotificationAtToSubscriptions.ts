import { MigrationInterface, QueryRunner, TableColumn } from 'typeorm';

export class AddLastNotificationAtToSubscriptions1776600000000 implements MigrationInterface {
  public async up(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.addColumn(
      'outlook_webhook_subscriptions',
      new TableColumn({
        name: 'last_notification_at',
        type: 'datetime',
        isNullable: true,
      })
    );
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.dropColumn(
      'outlook_webhook_subscriptions',
      'last_notification_at'
    );
  }
}
