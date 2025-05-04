import { Entity, Column, PrimaryGeneratedColumn, Index, CreateDateColumn } from 'typeorm';

/**
 * Entity for storing CSRF tokens for Microsoft OAuth
 * These tokens are used to protect against CSRF attacks
 * during the OAuth flow and are one-time use.
 */
@Entity('microsoft_csrf_tokens')
export class MicrosoftCsrfToken {
  @PrimaryGeneratedColumn()
  id: number = 0;

  @Column({ length: 64 })
  @Index({ unique: true })
  token: string = '';

  @Column({ name: 'user_id' })
  userId: string = '';

  @Column()
  expires: Date = new Date();

  @CreateDateColumn({ name: 'created_at' })
  createdAt: Date = new Date();
}
