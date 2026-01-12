import { Module } from '@nestjs/common';
import { TypeOrmModule } from '@nestjs/typeorm';
import { DeltaSyncService } from './delta-sync.service';
import { UserIdConverterService } from './user-id-converter.service';
import { MicrosoftUser } from '../../entities/microsoft-user.entity';
import { OutlookDeltaLink } from '../../entities/delta-link.entity';
import { OutlookDeltaLinkRepository } from '../../repositories/outlook-delta-link.repository';

@Module({
  imports: [TypeOrmModule.forFeature([MicrosoftUser, OutlookDeltaLink])],
  providers: [DeltaSyncService, UserIdConverterService, OutlookDeltaLinkRepository],
  exports: [DeltaSyncService, UserIdConverterService, OutlookDeltaLinkRepository],
})
export class SharedModule {} 