import { MigrationInterface, QueryRunner, TableColumn } from 'typeorm';

export class UpdateUserCalendarEntity1699000000001 implements MigrationInterface {
  public async up(queryRunner: QueryRunner): Promise<void> {
    // Check if table exists
    const table = await queryRunner.getTable('user_calendars');
    
    if (table) {
      // Remove token-related columns from user_calendars
      const tokenColumns = ['access_token', 'refresh_token', 'token_expiry'];
      
      for (const columnName of tokenColumns) {
        const column = table.findColumnByName(columnName);
        if (column) {
          await queryRunner.dropColumn('user_calendars', columnName);
          console.log(`Dropped column ${columnName} from user_calendars table`);
        }
      }
      
      // Add externalUserId column if it doesn't exist
      const externalUserIdColumn = table.findColumnByName('external_user_id');
      if (!externalUserIdColumn) {
        await queryRunner.addColumn(
          'user_calendars',
          new TableColumn({
            name: 'external_user_id',
            type: 'varchar',
            length: '255',
            isNullable: true, // Making nullable for existing data
          })
        );
        
        // Update existing records to set external_user_id to the same as user_id (as string)
        await queryRunner.query(`
          UPDATE user_calendars
          SET external_user_id = CAST(user_id AS TEXT)
          WHERE external_user_id IS NULL
        `);
        
        // Now make it non-nullable since we've populated data
        await queryRunner.changeColumn(
          'user_calendars',
          'external_user_id',
          new TableColumn({
            name: 'external_user_id',
            type: 'varchar',
            length: '255',
            isNullable: false,
          })
        );
      }
    }
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    // Check if table exists
    const table = await queryRunner.getTable('user_calendars');
    
    if (table) {
      // Add back the token columns
      const accessTokenColumn = table.findColumnByName('access_token');
      if (!accessTokenColumn) {
        await queryRunner.addColumn(
          'user_calendars',
          new TableColumn({
            name: 'access_token',
            type: 'text',
            isNullable: true,
          })
        );
      }
      
      const refreshTokenColumn = table.findColumnByName('refresh_token');
      if (!refreshTokenColumn) {
        await queryRunner.addColumn(
          'user_calendars',
          new TableColumn({
            name: 'refresh_token',
            type: 'text',
            isNullable: true,
          })
        );
      }
      
      const tokenExpiryColumn = table.findColumnByName('token_expiry');
      if (!tokenExpiryColumn) {
        await queryRunner.addColumn(
          'user_calendars',
          new TableColumn({
            name: 'token_expiry',
            type: 'datetime',
            isNullable: true,
          })
        );
      }
      
      // Remove external_user_id column if it was added
      const externalUserIdColumn = table.findColumnByName('external_user_id');
      if (externalUserIdColumn) {
        await queryRunner.dropColumn('user_calendars', 'external_user_id');
      }
    }
  }
} 