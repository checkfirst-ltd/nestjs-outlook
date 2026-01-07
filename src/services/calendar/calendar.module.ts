import { Module } from '@nestjs/common';
import { DeltaSyncService } from '../shared/delta-sync.service';
import { LifecycleEventHandlerService } from './lifecycle-event-handler.service';

@Module({
  providers: [
    DeltaSyncService,
    LifecycleEventHandlerService,
  ],
  exports: [
    DeltaSyncService,
    LifecycleEventHandlerService,
  ],
})
export class CalendarModule {} 