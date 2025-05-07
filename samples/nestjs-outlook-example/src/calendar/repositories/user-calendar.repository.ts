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

  async saveCalendarCredentials(
    userId: number,
    calendarId: string,
    accessToken: string,
    refreshToken: string,
    tokenExpiry: Date,
  ): Promise<UserCalendar> {
    const calendar = this.create({
      userId,
      calendarId,
      accessToken,
      refreshToken,
      tokenExpiry,
    });

    return this.save(calendar);
  }

  async updateTokens(
    id: number,
    accessToken: string,
    refreshToken: string,
    tokenExpiry: Date,
  ): Promise<UserCalendar | null> {
    await this.update(id, {
      accessToken,
      refreshToken,
      tokenExpiry,
    });

    return this.findOne({ where: { id } });
  }
} 