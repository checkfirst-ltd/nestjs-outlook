import { MigrationInterface, QueryRunner, Table, TableIndex, TableForeignKey } from "typeorm";

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

        // Add foreign key constraint
        await queryRunner.createForeignKey(
            'outlook_delta_links',
            new TableForeignKey({
                name: 'FK_outlook_delta_links_user',
                columnNames: ['user_id'],
                referencedColumnNames: ['id'],
                referencedTableName: 'microsoft_users',
                onDelete: 'CASCADE',
            })
        );
    }

    public async down(queryRunner: QueryRunner): Promise<void> {
        const table = await queryRunner.getTable('outlook_delta_links');
        const foreignKey = table?.foreignKeys.find(fk => fk.name === 'FK_outlook_delta_links_user');
        if (foreignKey) {
            await queryRunner.dropForeignKey('outlook_delta_links', foreignKey);
        }
        await queryRunner.dropIndex('outlook_delta_links', 'IDX_outlook_delta_links_user_id');
        await queryRunner.dropTable('outlook_delta_links');
    }
}