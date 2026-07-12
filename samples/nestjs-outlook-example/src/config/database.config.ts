import * as path from 'path';
import { TypeOrmModuleOptions } from '@nestjs/typeorm';
import { DataSource, DataSourceOptions } from 'typeorm';

// Resolve the path to the nestjs-outlook package
const outlookPackagePath = path.dirname(require.resolve('@checkfirst/nestjs-outlook/package.json'));

// Common TypeORM configuration used for both NestJS module and CLI
export const databaseConfig: DataSourceOptions = {
  type: 'sqlite',
  database: 'db.sqlite',
  entities: [
    // App entities
    __dirname + '/../**/*.entity{.ts,.js}',
    // Outlook module entities
    path.join(outlookPackagePath, 'dist', 'entities', '*.entity.js'),
  ],
  migrations: [
    // Local migrations in the sample app
    __dirname + '/../migrations/**/*.ts',
  ],
  synchronize: false, // Don't use synchronize in production
  logging: ['error', 'warn', 'schema'],
};

// Configuration for NestJS TypeORM module
export const typeOrmModuleOptions: TypeOrmModuleOptions = {
  ...databaseConfig,
  autoLoadEntities: true,
  migrationsRun: true,
};

// Create and export a data source for TypeORM CLI
export default new DataSource(databaseConfig); 