import {
    Entity,
    PrimaryGeneratedColumn,
    Column,
    CreateDateColumn,
    UpdateDateColumn,
    Index,
    ManyToOne,
    JoinColumn,
  } from 'typeorm';
  import { ResourceType } from '../enums/resource-type.enum';
  import { MicrosoftUser } from './microsoft-user.entity';
  
  @Entity('outlook_delta_links')
  export class OutlookDeltaLink {
    @PrimaryGeneratedColumn('increment')
    id!: number;
  
    @ManyToOne(() => MicrosoftUser)
    @JoinColumn({ name: 'user_id' })
    user!: MicrosoftUser;
  
    @Column({ name: 'user_id' })
    @Index()
    userId!: number;
  
    @Column({
      type: 'text',
      name: 'resource_type',
    })
    resourceType: ResourceType = ResourceType.CALENDAR;
  
    @Column({ name: 'delta_link', type: 'text' })
    deltaLink: string = '';
  
    @CreateDateColumn({ name: 'created_at' })
    createdAt: Date = new Date();
  
    @UpdateDateColumn({ name: 'updated_at' })
    updatedAt: Date = new Date();
  }