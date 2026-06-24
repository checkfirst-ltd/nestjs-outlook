import { Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { Repository } from 'typeorm';
import { MicrosoftTenantUser } from '../entities/microsoft-tenant-user.entity';
import { TtlCache } from '../utils/ttl-cache.util';

@Injectable()
export class MicrosoftTenantUserRepository {
  /**
   * Cache by external user ID: key = `${tenantId}:ext:${externalUserId}`
   */
  private readonly byExternalUserId = new TtlCache<string, MicrosoftTenantUser>(60000);

  /**
   * Cache by Microsoft user ID: key = `${tenantId}:ms:${microsoftUserId}`
   */
  private readonly byMicrosoftUserId = new TtlCache<string, MicrosoftTenantUser>(60000);

  constructor(
    @InjectRepository(MicrosoftTenantUser)
    private readonly repository: Repository<MicrosoftTenantUser>,
  ) {}

  private getCacheKeyExternal(tenantId: number, externalUserId: string): string {
    return `${tenantId}:ext:${externalUserId}`;
  }

  private getCacheKeyMicrosoft(tenantId: number, microsoftUserId: string): string {
    return `${tenantId}:ms:${microsoftUserId}`;
  }

  private invalidate(tenantUser?: Partial<MicrosoftTenantUser> | null, tenantId?: number): void {
    if (!tenantUser) return;

    const tenantObj = tenantUser.tenant as { id?: number } | null | undefined;
    const tid = tenantId ?? (tenantObj ? tenantObj.id : undefined);
    if (tid === undefined) return;

    if (tenantUser.externalUserId) {
      this.byExternalUserId.delete(this.getCacheKeyExternal(tid, tenantUser.externalUserId));
    }
    if (tenantUser.microsoftUserId) {
      this.byMicrosoftUserId.delete(this.getCacheKeyMicrosoft(tid, tenantUser.microsoftUserId));
    }
  }

  /**
   * Save a tenant user (create or update).
   */
  async save(tenantUser: Partial<MicrosoftTenantUser>): Promise<MicrosoftTenantUser> {
    // Get tenant ID from the relation or existing record
    const tenantObj = tenantUser.tenant as { id?: number } | null | undefined;
    const tenantId = tenantObj ? tenantObj.id : undefined;

    // Check if a user with this external ID already exists for the tenant
    if (tenantId && tenantUser.externalUserId) {
      const existing = await this.repository.findOne({
        where: {
          tenant: { id: tenantId },
          externalUserId: tenantUser.externalUserId,
        },
        relations: ['tenant'],
      });

      if (existing) {
        const originalId = existing.id;
        Object.assign(existing, tenantUser);
        existing.id = originalId;
        const saved = await this.repository.save(existing);
        this.invalidate(saved, tenantId);
        return saved;
      }
    }

    const userWithoutId = { ...tenantUser };
    delete userWithoutId.id;
    const newUser = this.repository.create(userWithoutId);

    const saved = await this.repository.save(newUser);
    this.invalidate(saved, tenantId);
    return saved;
  }

  /**
   * Find a tenant user by external user ID (host application's user ID).
   */
  async findByExternalUserId(
    tenantId: number,
    externalUserId: string,
  ): Promise<MicrosoftTenantUser | null> {
    const cacheKey = this.getCacheKeyExternal(tenantId, externalUserId);
    const cached = this.byExternalUserId.get(cacheKey);
    if (cached !== undefined) return cached;

    const result = await this.repository.findOne({
      where: {
        tenant: { id: tenantId },
        externalUserId,
        isActive: true,
      },
      relations: ['tenant'],
    });

    if (result) {
      this.byExternalUserId.set(cacheKey, result);
      // Also cache by Microsoft user ID
      const msKey = this.getCacheKeyMicrosoft(tenantId, result.microsoftUserId);
      this.byMicrosoftUserId.set(msKey, result);
    }

    return result;
  }

  /**
   * Find a tenant user by Microsoft user ID (Azure AD object ID).
   */
  async findByMicrosoftUserId(
    tenantId: number,
    microsoftUserId: string,
  ): Promise<MicrosoftTenantUser | null> {
    const cacheKey = this.getCacheKeyMicrosoft(tenantId, microsoftUserId);
    const cached = this.byMicrosoftUserId.get(cacheKey);
    if (cached !== undefined) return cached;

    const result = await this.repository.findOne({
      where: {
        tenant: { id: tenantId },
        microsoftUserId,
        isActive: true,
      },
      relations: ['tenant'],
    });

    if (result) {
      this.byMicrosoftUserId.set(cacheKey, result);
      // Also cache by external user ID
      const extKey = this.getCacheKeyExternal(tenantId, result.externalUserId);
      this.byExternalUserId.set(extKey, result);
    }

    return result;
  }

  /**
   * Find all users for a tenant.
   */
  async findAllByTenantId(tenantId: number): Promise<MicrosoftTenantUser[]> {
    return this.repository.find({
      where: {
        tenant: { id: tenantId },
        isActive: true,
      },
      relations: ['tenant'],
    });
  }

  /**
   * Deactivate a tenant user by external user ID.
   */
  async deactivate(tenantId: number, externalUserId: string): Promise<void> {
    const user = await this.findByExternalUserId(tenantId, externalUserId);
    if (user) {
      await this.repository.update(
        { id: user.id },
        { isActive: false, updatedAt: new Date() },
      );
      this.invalidate(user, tenantId);
    }
  }

  /**
   * Update the default calendar ID for a user.
   */
  async updateDefaultCalendarId(
    tenantId: number,
    externalUserId: string,
    defaultCalendarId: string,
  ): Promise<void> {
    const user = await this.findByExternalUserId(tenantId, externalUserId);
    if (user) {
      await this.repository.update(
        { id: user.id },
        { defaultCalendarId, updatedAt: new Date() },
      );
      this.invalidate(user, tenantId);
    }
  }

  /**
   * Delete a tenant user permanently.
   */
  async delete(tenantId: number, externalUserId: string): Promise<void> {
    const user = await this.findByExternalUserId(tenantId, externalUserId);
    if (user) {
      this.invalidate(user, tenantId);
      await this.repository.delete({ id: user.id });
    }
  }

  /**
   * Count active users for a tenant.
   */
  async countByTenantId(tenantId: number): Promise<number> {
    return this.repository.count({
      where: {
        tenant: { id: tenantId },
        isActive: true,
      },
    });
  }

  /**
   * Clear all caches.
   */
  clearCache(): void {
    this.byExternalUserId.clear();
    this.byMicrosoftUserId.clear();
  }
}
