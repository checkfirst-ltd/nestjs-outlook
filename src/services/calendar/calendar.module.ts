import { Module } from '@nestjs/common';
import { LifecycleEventHandlerService } from './lifecycle-event-handler.service';
import { SharedModule } from '../shared/shared.module';

@Module({
  imports: [SharedModule],
  providers: [
    LifecycleEventHandlerService,
  ],
  exports: [
    LifecycleEventHandlerService,
  ],
})
export class CalendarModule {} 