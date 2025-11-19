import { MigrationInterface, QueryRunner, Table, TableIndex } from "typeorm";

export class CreateOutlookTables1697025846000 implements MigrationInterface {
  public async up(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.createTable(
      new Table({
        name: "outlook_webhook_subscriptions",
        columns: [
          {
            name: "id",
            type: "integer",
            isPrimary: true,
            isGenerated: true,
            generationStrategy: "increment",
          },
          {
            name: "subscription_id",
            type: "varchar",
            length: "255",
            isNullable: false,
            isUnique: true,
          },
          {
            name: "user_id",
            type: "integer",
            isNullable: false,
          },
          {
            name: "resource",
            type: "varchar",
            length: "255",
            isNullable: false,
          },
          {
            name: "change_type",
            type: "varchar",
            length: "255",
            isNullable: false,
          },
          {
            name: "client_state",
            type: "varchar",
            length: "255",
            isNullable: false,
          },
          {
            name: "notification_url",
            type: "varchar",
            length: "255",
            isNullable: false,
          },
          {
            name: "expiration_date_time",
            type: "timestamp",
            isNullable: false,
          },
          {
            name: "is_active",
            type: "boolean",
            isNullable: false,
            default: true,
          },
          {
            name: "access_token",
            type: "text",
            isNullable: false,
          },
          {
            name: "refresh_token",
            type: "text",
            isNullable: false,
          },
          {
            name: "created_at",
            type: "timestamp",
            isNullable: false,
            default: "NOW()",
          },
          {
            name: "updated_at",
            type: "timestamp",
            isNullable: false,
            default: "NOW()",
          },
        ],
      }),
      true
    );

    await queryRunner.createIndex(
      "outlook_webhook_subscriptions",
      new TableIndex({
        name: "UQ_outlook_webhook_subscriptions_id",
        columnNames: ["subscription_id"],
        isUnique: true,
      })
    );

    await queryRunner.createTable(
      new Table({
        name: "microsoft_csrf_tokens",
        columns: [
          {
            name: "id",
            type: "integer",
            isPrimary: true,
            isGenerated: true,
            generationStrategy: "increment",
          },
          {
            name: "token",
            type: "varchar",
            length: "64",
            isNullable: false,
          },
          {
            name: "user_id",
            type: "varchar",
            length: "255",
            isNullable: false,
          },
          {
            name: "expires",
            type: "timestamp",
            isNullable: false,
          },
          {
            name: "created_at",
            type: "timestamp",
            isNullable: false,
            default: "NOW()",
          },
        ],
      }),
      true
    );

    await queryRunner.createIndex(
      "microsoft_csrf_tokens",
      new TableIndex({
        name: "UQ_microsoft_csrf_tokens_token",
        columnNames: ["token"],
        isUnique: true,
      })
    );
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.query(
      `DROP TABLE IF EXISTS outlook_webhook_subscriptions`
    );
    await queryRunner.query(`DROP TABLE IF EXISTS microsoft_csrf_tokens`);
  }
}
