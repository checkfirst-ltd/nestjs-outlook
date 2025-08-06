import { DataSource } from 'typeorm';
import { config } from 'dotenv';
import * as path from 'path';

// Load environment variables
config();

export default new DataSource({
  type: 'postgres',
  host: process.env.DB_HOST || 'localhost',
  port: parseInt(process.env.DB_PORT || '5432'),
  username: process.env.DB_USERNAME || 'postgres',
  password: process.env.DB_PASSWORD || 'postgres',
  database: process.env.DB_DATABASE || 'nestjs_outlook',
  entities: [
    path.join(__dirname, '..', '**', '*.entity{.ts,.js}'),
    path.join(__dirname, '..', '..', '..', 'dist', 'entities', '*.entity.js')
  ],
  migrations: [path.join(__dirname, '..', 'migrations', '*.ts')],
  synchronize: false,
}); 