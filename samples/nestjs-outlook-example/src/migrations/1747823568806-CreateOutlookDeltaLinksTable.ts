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
                        name: 'user_id',
                        type: 'integer',
                        isNullable: false,
                    },
                    {
                        name: 'resource_type',
                        type: 'text',
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

        // Create index on user_id
        await queryRunner.createIndex(
            'outlook_delta_links',
            new TableIndex({
                name: 'IDX_outlook_delta_links_user_id',
                columnNames: ['user_id'],
            })
        );


    }

    public async down(queryRunner: QueryRunner): Promise<void> {
        await queryRunner.dropIndex('outlook_delta_links', 'IDX_outlook_delta_links_user_id');
        await queryRunner.dropTable('outlook_delta_links');
    }
}