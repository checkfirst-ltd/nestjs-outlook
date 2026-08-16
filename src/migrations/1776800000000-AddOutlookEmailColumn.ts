import { MigrationInterface, QueryRunner, TableColumn } from 'typeorm';

export class AddOutlookEmailColumn1776800000000 implements MigrationInterface {
  public async up(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.addColumn(
      'microsoft_users',
      new TableColumn({
        name: 'outlook_email',
        type: 'varchar',
        length: '320',
        isNullable: true,
      })
    );
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.dropColumn('microsoft_users', 'outlook_email');
  }
}
