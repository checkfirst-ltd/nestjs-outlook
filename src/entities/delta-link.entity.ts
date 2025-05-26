import {
    Entity,
    PrimaryGeneratedColumn,
    Column,
    CreateDateColumn,
    UpdateDateColumn,
    Index,
  } from 'typeorm';
  
  @Entity('outlook_delta_links')
  export class OutlookDeltaLink {
    @PrimaryGeneratedColumn('increment')
    id!: number;
  
    @Column({ name: 'external_user_id' })
    @Index()
    externalUserId: string = '';
  
    @Column({ name: 'resource_type' })
    resourceType: string = ''; // 'calendar', 'email'
  
    @Column({ name: 'delta_link', type: 'text' })
    deltaLink: string = '';
  
    @CreateDateColumn({ name: 'created_at' })
    createdAt: Date = new Date();
  
    @UpdateDateColumn({ name: 'updated_at' })
    updatedAt: Date = new Date();
  }