import { MigrationInterface, QueryRunner, TableColumn } from 'typeorm';

export class AddMicrosoftUserStatusColumn1776700000000 implements MigrationInterface {
  public async up(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.addColumn(
      'microsoft_users',
      new TableColumn({
        name: 'status',
        type: 'varchar',
        length: '32',
        isNullable: false,
        default: "'ACTIVE'",
      })
    );
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.dropColumn('microsoft_users', 'status');
  }
}
