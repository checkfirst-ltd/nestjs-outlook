import { MigrationInterface, QueryRunner, Table, TableIndex } from "typeorm";

export class CreateOutlookDeltaLinksTable1747823568806 implements MigrationInterface {
    name = 'CreateOutlookDeltaLinksTable1747823568806'

    public async up(queryRunner: QueryRunner): Promise<void> {
        await queryRunner.createTable(
            new Table({
                name: 'outlook_delta_links',
                columns: [
                    {
                        name: 'id',
                        type: 'integer',
                        isPrimary: true,
                        isGenerated: true,
                        generationStrategy: 'increment',
                    },
                    {
                        name: 'external_user_id',
                        type: 'varchar',
                        isNullable: false,
                    },
                    {
                        name: 'resource_type',
                        type: 'varchar',
                        isNullable: false,
                    },
                    {
                        name: 'delta_link',
                        type: 'text',
                        isNullable: false,
                    },
                    {
                        name: 'created_at',
                        type: 'datetime',
                        default: "datetime('now')",
                        isNullable: false,
                    },
                    {
                        name: 'updated_at',
                        type: 'datetime',
                        default: "datetime('now')",
                        isNullable: false,
                    },
                ],
            }),
            true
        );
        await queryRunner.createIndex(
            'outlook_delta_links',
            new TableIndex({
                name: 'IDX_045ba046da3745437f2b3f8903',
                columnNames: ['external_user_id'],
            })
        );
    }

    public async down(queryRunner: QueryRunner): Promise<void> {
        await queryRunner.dropIndex('outlook_delta_links', 'IDX_045ba046da3745437f2b3f8903');
        await queryRunner.dropTable('outlook_delta_links');
    }
}