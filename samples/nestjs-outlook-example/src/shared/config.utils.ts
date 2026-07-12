import { ConfigService } from '@nestjs/config';

export function getRequiredConfig(configService: ConfigService, key: string): string {
  const value = configService.get<string>(key);
  if (!value) {
    throw new Error(`Required environment variable ${key} is not set`);
  }
  return value;
} 