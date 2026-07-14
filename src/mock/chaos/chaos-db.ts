import { FindOperator } from 'typeorm';
import { MicrosoftUser } from '../../entities/microsoft-user.entity';
import { MicrosoftTenant } from '../../entities/microsoft-tenant.entity';
import { MicrosoftTenantStatus } from '../../enums/microsoft-tenant-status.enum';
import { OutlookWebhookSubscription } from '../../entities/outlook-webhook-subscription.entity';
import { ChaosEngine, chaosDelay } from './chaos-engine';
import { ChaosMetrics } from './chaos-metrics';

/**
 * In-memory database behind the repository fakes.
 *
 * The fakes mirror the REAL repositories' query semantics (active flags, `expirationDateTime >
 * now` filters, upsert-by-subscriptionId, the query-builder shapes used by
 * `clearTenantUserMappings`) so chaos tests exercise genuine behaviour, not mock behaviour.
 * Every method can be disrupted through the shared {@link ChaosEngine} using `db.<method>`
 * routes, plus a table-wide latency range (virtual time under fake timers).
 */
export class ChaosDb {
  readonly users: MicrosoftUser[] = [];
  readonly tenants: MicrosoftTenant[] = [];
  readonly subscriptions: OutlookWebhookSubscription[] = [];
  private userSeq = 0;
  private tenantSeq = 0;
  private subSeq = 0;

  constructor(
    private readonly engine: ChaosEngine,
    private readonly metrics: ChaosMetrics,
    private readonly latencyMs: { min: number; max: number } = { min: 0, max: 0 },
  ) {}

  /** Latency + planned-failure hook wrapped around every repo method. */
  private async touch(method: string, key: string): Promise<void> {
    this.metrics.recordDb(method);
    if (this.latencyMs.max > 0) {
      await chaosDelay(this.engine.random.int(this.latencyMs.min, this.latencyMs.max));
    }
    const injected = this.engine.decide(`db.${method}`, key, { plansOnly: true });
    if (injected) {
      this.metrics.recordInjected(`db.${method}`, injected.response?.status ?? 'network');
      throw new Error(`chaos db failure on ${method} (${key})`);
    }
  }

  // ── seeding ─────────────────────────────────────────────────────────

  addTenant(partial: Partial<MicrosoftTenant>): MicrosoftTenant {
    this.tenantSeq += 1;
    const tenant = Object.assign(new MicrosoftTenant(), { id: this.tenantSeq, isActive: true }, partial);
    this.tenants.push(tenant);
    return tenant;
  }

  addUser(partial: Partial<MicrosoftUser>): MicrosoftUser {
    this.userSeq += 1;
    const user = Object.assign(new MicrosoftUser(), { id: this.userSeq }, partial);
    this.users.push(user);
    return user;
  }

  addSubscription(partial: Partial<OutlookWebhookSubscription>): OutlookWebhookSubscription {
    this.subSeq += 1;
    const sub = Object.assign(new OutlookWebhookSubscription(), { id: this.subSeq, isActive: true }, partial);
    this.subscriptions.push(sub);
    return sub;
  }

  activeSubsOfUser(userId: number): OutlookWebhookSubscription[] {
    return this.subscriptions.filter((s) => s.userId === userId && s.isActive);
  }

  // ── fakes ───────────────────────────────────────────────────────────

  /** Structural stand-in for `OutlookWebhookSubscriptionRepository` (real query semantics). */
  buildSubscriptionRepo() {
    const now = (): number => Date.now();
    return {
      saveSubscription: async (partial: Partial<OutlookWebhookSubscription>) => {
        // Chaos key by externalUserId so tests can plan failures deterministically
        // (internal ids depend on async creation order under concurrency).
        const owner = this.users.find((u) => u.id === partial.userId);
        await this.touch('subs.save', owner?.externalUserId ?? String(partial.userId ?? '?'));
        const existing = this.subscriptions.find((s) => s.subscriptionId === partial.subscriptionId);
        if (existing) {
          Object.assign(existing, partial, { id: existing.id });
          return existing;
        }
        return this.addSubscription(partial);
      },
      findAllActiveByUserId: async (userId: number) => {
        await this.touch('subs.findAllActiveByUserId', String(userId));
        return this.subscriptions.filter((s) => s.userId === userId && s.isActive);
      },
      findActiveByUserId: async (userId: number) => {
        await this.touch('subs.findActiveByUserId', String(userId));
        return this.subscriptions.find((s) => s.userId === userId && s.isActive) ?? null;
      },
      findActiveByUserIds: async (userIds: number[]) => {
        await this.touch('subs.findActiveByUserIds', String(userIds.length));
        const ids = new Set(userIds);
        // Mirrors the real repository: active AND not expired.
        return this.subscriptions.filter(
          (s) => ids.has(s.userId) && s.isActive && s.expirationDateTime.getTime() > now(),
        );
      },
      findAllActiveByTenantId: async (tenantId: string) => {
        await this.touch('subs.findAllActiveByTenantId', tenantId);
        return this.subscriptions.filter(
          (s) => s.tenantId === tenantId && s.isActive && s.expirationDateTime.getTime() > now(),
        );
      },
      findActiveByTenantAndMicrosoftUser: async (tenantId: string, microsoftUserId: string) => {
        await this.touch('subs.findActiveByTenantAndMicrosoftUser', microsoftUserId);
        return (
          this.subscriptions.find(
            (s) =>
              s.tenantId === tenantId &&
              s.microsoftUserId === microsoftUserId &&
              s.isActive &&
              s.expirationDateTime.getTime() > now(),
          ) ?? null
        );
      },
      deactivateSubscription: async (subscriptionId: string) => {
        await this.touch('subs.deactivate', subscriptionId);
        const sub = this.subscriptions.find((s) => s.subscriptionId === subscriptionId);
        if (sub) sub.isActive = false;
      },
      deactivateAllByTenantId: async (tenantId: string) => {
        await this.touch('subs.deactivateAllByTenantId', tenantId);
        let affected = 0;
        for (const sub of this.subscriptions) {
          if (sub.tenantId === tenantId && sub.isActive) {
            sub.isActive = false;
            affected += 1;
          }
        }
        return affected;
      },
    };
  }

  /** Structural stand-in for the TypeORM `Repository<MicrosoftUser>` subset the services use. */
  buildUserOrmRepo() {
    const matches = (user: MicrosoftUser, where: Record<string, unknown>): boolean => {
      for (const [field, expected] of Object.entries(where)) {
        const actual = (user as unknown as Record<string, unknown>)[field];
        if (expected instanceof FindOperator) {
          if (expected.type === 'in') {
            const values = expected.value as unknown as unknown[];
            if (!values.includes(actual)) return false;
          } else {
            throw new Error(`ChaosDb: unsupported FindOperator ${expected.type}`);
          }
        } else if (actual !== expected) {
          return false;
        }
      }
      return true;
    };

    // Reads return SHALLOW CLONES, mirroring a real ORM: mutating a loaded entity must not
    // change the database until save() is called. (Handing out live table references would
    // silently persist mutations even when the save is chaos-failed.)
    const clone = (user: MicrosoftUser): MicrosoftUser => Object.assign(new MicrosoftUser(), user);

    return {
      find: async (options: { where: Record<string, unknown> }) => {
        await this.touch('users.find', JSON.stringify(Object.keys(options.where)));
        return this.users.filter((u) => matches(u, options.where)).map(clone);
      },
      findOne: async (options: { where: Record<string, unknown> }) => {
        await this.touch('users.findOne', String(options.where.externalUserId ?? options.where.id ?? '?'));
        const row = this.users.find((u) => matches(u, options.where));
        return row ? clone(row) : null;
      },
      save: async (user: MicrosoftUser) => {
        await this.touch('users.save', user.externalUserId);
        const existing = this.users.find((u) => u.externalUserId === user.externalUserId);
        if (existing) {
          Object.assign(existing, user, { id: existing.id });
          return clone(existing);
        }
        return clone(this.addUser(Object.assign(new MicrosoftUser(), user)));
      },
      createQueryBuilder: (_alias?: string) => this.buildUserQueryBuilder(),
    };
  }

  /**
   * Interprets the exact query-builder shapes `clearTenantUserMappings` issues:
   * select refresh tokens / bulk unmap UPDATE / bulk DELETE, filtered by `tenant_id = :id`
   * and `refresh_token IS [NOT] NULL`.
   */
  private buildUserQueryBuilder() {
    let tenantInternalId: number | null = null;
    let refreshTokenFilter: 'not-null' | 'null' | null = null;
    let operation: 'select' | 'update' | 'delete' = 'select';
    let setPayload: Partial<MicrosoftUser> | null = null;

    const applyWhere = (clause: string, params?: Record<string, unknown>): void => {
      if (clause.includes('tenant_id = :id')) tenantInternalId = Number(params?.id);
      else if (clause.includes('refresh_token IS NOT NULL')) refreshTokenFilter = 'not-null';
      else if (clause.includes('refresh_token IS NULL')) refreshTokenFilter = 'null';
      else throw new Error(`ChaosDb: unsupported where clause "${clause}"`);
    };

    const selectRows = (): MicrosoftUser[] =>
      this.users.filter((u) => {
        if (u.tenant?.id !== tenantInternalId) return false;
        if (refreshTokenFilter === 'not-null') return u.refreshToken !== null;
        if (refreshTokenFilter === 'null') return u.refreshToken === null;
        return true;
      });

    const qb = {
      select: (_field: string, _aliasName: string) => qb,
      where: (clause: string, params?: Record<string, unknown>) => {
        applyWhere(clause, params);
        return qb;
      },
      andWhere: (clause: string, params?: Record<string, unknown>) => {
        applyWhere(clause, params);
        return qb;
      },
      update: () => {
        operation = 'update';
        return qb;
      },
      delete: () => {
        operation = 'delete';
        return qb;
      },
      set: (payload: Partial<MicrosoftUser>) => {
        setPayload = payload;
        return qb;
      },
      getRawMany: async <T>(): Promise<T[]> => {
        await this.touch('users.qb.select', String(tenantInternalId));
        return selectRows().map((u) => ({ refreshToken: u.refreshToken }) as T);
      },
      execute: async (): Promise<{ affected: number }> => {
        await this.touch(`users.qb.${operation}`, String(tenantInternalId));
        const rows = selectRows();
        if (operation === 'update' && setPayload) {
          for (const row of rows) Object.assign(row, setPayload);
        } else if (operation === 'delete') {
          for (const row of rows) this.users.splice(this.users.indexOf(row), 1);
        }
        return { affected: rows.length };
      },
    };
    return qb;
  }

  /** Structural stand-in for the TypeORM `Repository<MicrosoftTenant>` subset. */
  buildTenantOrmRepo() {
    return {
      findOne: async (options: { where: { tenantId: string; isActive?: boolean } }) => {
        await this.touch('tenants.findOne', options.where.tenantId);
        return (
          this.tenants.find(
            (t) =>
              t.tenantId === options.where.tenantId &&
              (options.where.isActive === undefined || t.isActive === options.where.isActive),
          ) ?? null
        );
      },
    };
  }

  /** Structural stand-in for the custom `MicrosoftTenantRepository` (connection repo). */
  buildTenantConnectionRepo() {
    return {
      findByTenantId: async (tenantId: string) => {
        await this.touch('tenants.findByTenantId', tenantId);
        // Mirrors the real repository: only active connections are returned.
        return this.tenants.find((t) => t.tenantId === tenantId && t.isActive) ?? null;
      },
      findAllActive: async () => {
        await this.touch('tenants.findAllActive', '*');
        // Mirrors the real repository: isActive AND status === ACTIVE.
        return this.tenants.filter((t) => t.isActive && t.status === MicrosoftTenantStatus.ACTIVE);
      },
      deactivate: async (tenantId: string) => {
        await this.touch('tenants.deactivate', tenantId);
        const tenant = this.tenants.find((t) => t.tenantId === tenantId);
        if (tenant) tenant.isActive = false;
        this.metrics.mark('tenant:deactivate');
      },
    };
  }
}
