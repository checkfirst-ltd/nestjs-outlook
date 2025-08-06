import { Module } from '@nestjs/common';
import { DeltaSyncService } from './delta-sync.service';

@Module({
  providers: [DeltaSyncService],
  exports: [DeltaSyncService],
})
export class SharedModule {} 