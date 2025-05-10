import { Injectable } from '@nestjs/common';
import { DataSource, Repository } from 'typeorm';
import { UserCalendar } from '../entities/user-calendar.entity';

@Injectable()
export class UserCalendarRepository extends Repository<UserCalendar> {
  constructor(private dataSource: DataSource) {
    super(UserCalendar, dataSource.createEntityManager());
  }

  async findByUserId(userId: number): Promise<UserCalendar[]> {
    return this.find({
      where: {
        userId,
        active: true,
      },
    });
  }

  async findActiveByUserId(userId: number): Promise<UserCalendar | null> {
    return this.findOne({
      where: {
        userId,
        active: true,
      },
    });
  }

  async saveCalendarDetails(
    userId: number,
    externalUserId: string,
    calendarId: string,
  ): Promise<UserCalendar> {
    const calendar = this.create({
      userId,
      externalUserId,
      calendarId,
      active: true,
    });

    return this.save(calendar);
  }
} 