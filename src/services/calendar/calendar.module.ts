import { Module } from '@nestjs/common';
import { DeltaSyncService } from '../shared/delta-sync.service';

@Module({
  providers: [
    DeltaSyncService,
  ],
  exports: [
    DeltaSyncService,
  ],
})
export class CalendarModule {} 