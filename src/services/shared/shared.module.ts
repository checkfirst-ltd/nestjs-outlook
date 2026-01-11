import { Module } from '@nestjs/common';
import { TypeOrmModule } from '@nestjs/typeorm';
import { DeltaSyncService } from './delta-sync.service';
import { UserIdConverterService } from './user-id-converter.service';
import { MicrosoftUser } from '../../entities/microsoft-user.entity';

@Module({
  imports: [TypeOrmModule.forFeature([MicrosoftUser])],
  providers: [DeltaSyncService, UserIdConverterService],
  exports: [DeltaSyncService, UserIdConverterService],
})
export class SharedModule {} 