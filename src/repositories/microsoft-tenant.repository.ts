import { Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { Repository } from 'typeorm';
import { MicrosoftTenant } from '../entities/microsoft-tenant.entity';
import { MicrosoftTenantStatus } from '../enums/microsoft-tenant-status.enum';
import { TtlCache } from '../utils/ttl-cache.util';

@Injectable()
export class MicrosoftTenantRepository {
  private readonly byTenantId = new TtlCache<string, MicrosoftTenant>(60000);

  constructor(
    @InjectRepository(MicrosoftTenant)
    private readonly repository: Repository<MicrosoftTenant>,
  ) {}

  private invalidate(tenant?: Partial<MicrosoftTenant> | null): void {
    if (tenant?.tenantId) {
      this.byTenantId.delete(tenant.tenantId);
    }
  }

  /**
   * Save a tenant (create or update).
   */
  async save(tenant: Partial<MicrosoftTenant>): Promise<MicrosoftTenant> {
    // Check if a tenant with this tenantId already exists
    if (tenant.tenantId) {
      const existing = await this.repository.findOne({
        where: { tenantId: tenant.tenantId },
      });

      if (existing) {
        const originalId = existing.id;
        Object.assign(existing, tenant);
        existing.id = originalId;
        const saved = await this.repository.save(existing);
        this.invalidate(saved);
        return saved;
      }
    }

    const tenantWithoutId = { ...tenant };
    delete tenantWithoutId.id;
    const newTenant = this.repository.create(tenantWithoutId);

    const saved = await this.repository.save(newTenant);
    this.invalidate(saved);
    return saved;
  }

  /**
   * Find a tenant by Microsoft tenant ID (Azure AD directory ID).
   */
  async findByTenantId(tenantId: string): Promise<MicrosoftTenant | null> {
    const cached = this.byTenantId.get(tenantId);
    if (cached !== undefined) return cached;

    const result = await this.repository.findOne({
      where: { tenantId, isActive: true },
    });
    if (result) this.byTenantId.set(tenantId, result);
    return result;
  }

  /**
   * Find a tenant by external tenant ID (state parameter in consent flow).
   * For this implementation, we use tenantId as the external identifier.
   */
  async findByExternalTenantId(externalTenantId: string): Promise<MicrosoftTenant | null> {
    // In this implementation, externalTenantId maps to tenantId
    return this.findByTenantId(externalTenantId);
  }

  /**
   * Find all active tenants.
   */
  async findAllActive(): Promise<MicrosoftTenant[]> {
    return this.repository.find({
      where: {
        isActive: true,
        status: MicrosoftTenantStatus.ACTIVE,
      },
    });
  }

  /**
   * Update the status of a tenant.
   */
  async updateStatus(
    tenantId: string,
    status: MicrosoftTenantStatus,
  ): Promise<void> {
    await this.repository.update(
      { tenantId },
      { status, updatedAt: new Date() },
    );
    this.byTenantId.delete(tenantId);
  }

  /**
   * Mark admin consent as granted.
   */
  async markConsentGranted(tenantId: string): Promise<void> {
    await this.repository.update(
      { tenantId },
      {
        status: MicrosoftTenantStatus.ACTIVE,
        adminConsentGrantedAt: new Date(),
        updatedAt: new Date(),
      },
    );
    this.byTenantId.delete(tenantId);
  }

  /**
   * Deactivate a tenant.
   */
  async deactivate(tenantId: string): Promise<void> {
    await this.repository.update(
      { tenantId },
      { isActive: false, updatedAt: new Date() },
    );
    this.byTenantId.delete(tenantId);
  }

  /**
   * Find tenants by status.
   */
  async findByStatus(status: MicrosoftTenantStatus): Promise<MicrosoftTenant[]> {
    return this.repository.find({
      where: { status, isActive: true },
    });
  }

  /**
   * Delete a tenant permanently.
   */
  async delete(tenantId: string): Promise<void> {
    const tenant = await this.repository.findOne({
      where: { tenantId },
    });
    if (tenant) {
      this.invalidate(tenant);
      await this.repository.delete({ tenantId });
    }
  }

  /**
   * Count tenants by status.
   */
  async countByStatus(status: MicrosoftTenantStatus): Promise<number> {
    return this.repository.count({
      where: { status, isActive: true },
    });
  }
}
