import { MigrationInterface, QueryRunner, TableColumn } from 'typeorm';

export class AddDefaultCalendarIdColumn1740400000000 implements MigrationInterface {
  public async up(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.addColumn(
      'microsoft_users',
      new TableColumn({
        name: 'default_calendar_id',
        type: 'varchar',
        length: '255',
        isNullable: true,
      })
    );
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.dropColumn('microsoft_users', 'default_calendar_id');
  }
}
