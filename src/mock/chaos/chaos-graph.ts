import { ChaosEngine, chaosDelay, buildChaosError } from './chaos-engine';
import { ChaosMetrics } from './chaos-metrics';

/** A Microsoft user living in the fake Graph directory. */
export interface GraphUser {
  id: string;
  userPrincipalName: string;
  displayName: string;
  mail: string | null;
}

/** A webhook subscription living in the fake Graph. */
export interface GraphSubscription {
  id: string;
  resource: string;
  changeType: string;
  clientState: string;
  notificationUrl: string;
  expirationDateTime: string;
}

interface GraphResponse {
  status: number;
  data: unknown;
}

interface BatchInnerRequest {
  id: string;
  method: string;
  url: string;
}

/** Minimal structural view of a jest-mocked axios module (avoids importing jest types). */
export interface AxiosMockLike {
  get: { mockImplementation: (impl: (url: string, config?: unknown) => Promise<unknown>) => unknown };
  post: {
    mockImplementation: (impl: (url: string, body?: unknown, config?: unknown) => Promise<unknown>) => unknown;
  };
  delete: { mockImplementation: (impl: (url: string, config?: unknown) => Promise<unknown>) => unknown };
  patch: {
    mockImplementation: (impl: (url: string, body?: unknown, config?: unknown) => Promise<unknown>) => unknown;
  };
  isAxiosError: { mockImplementation: (impl: (err: unknown) => boolean) => unknown };
}

/**
 * Chaos route names — targets for {@link ChaosEngine.alwaysFail}/`failTimes` plans and for
 * metrics assertions.
 *
 * - `users.lookup` (key = email) — GET /users?$filter
 * - `users.get` (key = upn) — GET /users/{upn}
 * - `subs.create` (key = msUserId or `me:{internalId}`) — POST /subscriptions
 * - `subs.delete` / `subs.get` (key = subscriptionId)
 * - `batch` (key = `*`, whole call) and `batch.item` (key = subscriptionId, per inner request)
 * - `auth.revoke` (key = refresh token) — POST …/logout
 */
export type GraphRoute =
  | 'users.lookup'
  | 'users.get'
  | 'subs.create'
  | 'subs.delete'
  | 'subs.get'
  | 'subs.list'
  | 'subs.patch'
  | 'batch'
  | 'batch.item'
  | 'auth.revoke';

/**
 * A stateful in-memory Microsoft Graph behind a chaos layer.
 *
 * The services under test call the (jest-mocked) `axios` module; `install()` routes those calls
 * here. Each attempt pays a sampled latency (virtual time under fake timers), may be disrupted
 * by the {@link ChaosEngine} (plans first, then random rates), and is recorded in
 * {@link ChaosMetrics} — so flows behave *functionally* like real Graph (create returns an id
 * that later GET/DELETE resolve; $batch fans out per item) while tests control the weather.
 */
export class ChaosGraph {
  readonly users = new Map<string, GraphUser>(); // keyed by lower-cased email/UPN
  readonly subscriptions = new Map<string, GraphSubscription>();
  private seq = 0;

  constructor(
    readonly engine: ChaosEngine,
    readonly metrics: ChaosMetrics,
  ) {}

  seedUser(email: string, msUserId: string): GraphUser {
    const user: GraphUser = {
      id: msUserId,
      userPrincipalName: email,
      displayName: email.split('@')[0],
      mail: email,
    };
    this.users.set(email.toLowerCase(), user);
    return user;
  }

  seedSubscription(sub: Omit<GraphSubscription, 'expirationDateTime'> & { expirationDateTime?: string }): void {
    this.subscriptions.set(sub.id, {
      ...sub,
      expirationDateTime: sub.expirationDateTime ?? new Date(Date.now() + 72 * 3600 * 1000).toISOString(),
    });
  }

  subscriptionIdsForResourcePrefix(prefix: string): string[] {
    return [...this.subscriptions.values()].filter((s) => s.resource.startsWith(prefix)).map((s) => s.id);
  }

  /**
   * Point of entry for every faked axios call. Resolves a response or throws an axios-shaped
   * error.
   *
   * Failure semantics: random-rate and default plan injections fire BEFORE `execute()` — the
   * request never reaches Graph (at-most-once). Plans registered via `failTimesAfterExecute`
   * fire AFTER `execute()` mutated state — the request took effect but the response was lost
   * (at-least-once), which is how a retried non-idempotent create duplicates a subscription.
   */
  private async dispatch(
    route: GraphRoute,
    key: string,
    execute: () => GraphResponse,
  ): Promise<GraphResponse> {
    this.metrics.enter(route, key);
    try {
      await chaosDelay(this.engine.latency());
      const decision = this.engine.decideFull(route, key);
      if (decision && !decision.afterExecute) {
        this.metrics.recordInjected(route, decision.error.response?.status ?? 'network');
        throw decision.error;
      }
      const result = execute();
      if (decision?.afterExecute) {
        // State IS mutated — the caller just never learns about it.
        this.metrics.recordInjected(route, decision.error.response?.status ?? 'network');
        throw decision.error;
      }
      if (result.status >= 400) {
        throw buildChaosError(result.status);
      }
      return result;
    } finally {
      this.metrics.exit();
    }
  }

  // ── route handlers ──────────────────────────────────────────────────

  private handleUsersLookup(params: Record<string, unknown> | undefined): Promise<GraphResponse> {
    const filter = String(params?.['$filter'] ?? '');
    const match = /'([^']*)'/.exec(filter);
    const email = (match?.[1] ?? '').toLowerCase();
    return this.dispatch('users.lookup', email, () => {
      const user = this.users.get(email);
      return { status: 200, data: { value: user ? [user] : [] } };
    });
  }

  private handleUsersGet(upn: string): Promise<GraphResponse> {
    const key = decodeURIComponent(upn).toLowerCase();
    return this.dispatch('users.get', key, () => {
      const user = this.users.get(key);
      return user ? { status: 200, data: user } : { status: 404, data: null };
    });
  }

  private handleSubsCreate(body: Record<string, unknown>): Promise<GraphResponse> {
    const resource = String(body.resource ?? '');
    const clientState = String(body.clientState ?? '');
    // Plan key: app-only → the Microsoft user id in the resource; delegated → `me:{internalId}`.
    const appOnlyMatch = /^\/users\/([^/]+)\/events$/.exec(resource);
    const delegatedMatch = /^user_(\d+)_/.exec(clientState);
    const key = appOnlyMatch ? appOnlyMatch[1] : `me:${delegatedMatch?.[1] ?? 'unknown'}`;

    return this.dispatch('subs.create', key, () => {
      this.seq += 1;
      const sub: GraphSubscription = {
        id: `graph-sub-${this.seq}`,
        resource,
        changeType: String(body.changeType ?? 'created,updated,deleted'),
        clientState,
        notificationUrl: String(body.notificationUrl ?? ''),
        expirationDateTime: String(body.expirationDateTime ?? new Date().toISOString()),
      };
      this.subscriptions.set(sub.id, sub);
      return { status: 201, data: sub };
    });
  }

  private handleSubsDelete(id: string): Promise<GraphResponse> {
    return this.dispatch('subs.delete', id, () => {
      if (!this.subscriptions.has(id)) return { status: 404, data: null };
      this.subscriptions.delete(id);
      return { status: 204, data: '' };
    });
  }

  private handleSubsGet(id: string): Promise<GraphResponse> {
    return this.dispatch('subs.get', id, () => {
      const sub = this.subscriptions.get(id);
      return sub ? { status: 200, data: sub } : { status: 404, data: null };
    });
  }

  private handleSubsList(): Promise<GraphResponse> {
    return this.dispatch('subs.list', '*', () => ({
      status: 200,
      data: { value: [...this.subscriptions.values()] },
    }));
  }

  private async handleBatch(body: { requests?: BatchInnerRequest[] }): Promise<GraphResponse> {
    // The outer $batch call is one HTTP request (chaos-able as a whole)…
    return this.dispatch('batch', '*', () => ({ status: 200, data: null })).then(async () => {
      // …then each inner request gets its own per-item chaos decision.
      const responses: { id: string; status: number; body: unknown }[] = [];
      for (const request of body.requests ?? []) {
        const idMatch = /\/subscriptions\/([^/?]+)/.exec(request.url);
        const subId = idMatch?.[1] ?? 'unknown';
        const injected = this.engine.decide('batch.item', subId);
        if (injected) {
          this.metrics.recordInjected('batch.item', injected.response?.status ?? 'network');
          responses.push({
            id: request.id,
            status: injected.response?.status ?? 500,
            body: injected.response?.data ?? null,
          });
          continue;
        }
        if (request.method === 'DELETE') {
          if (this.subscriptions.has(subId)) {
            this.subscriptions.delete(subId);
            responses.push({ id: request.id, status: 204, body: null });
          } else {
            responses.push({ id: request.id, status: 404, body: null });
          }
        } else if (request.method === 'GET') {
          const sub = this.subscriptions.get(subId);
          responses.push({ id: request.id, status: sub ? 200 : 404, body: sub ?? null });
        } else {
          responses.push({ id: request.id, status: 400, body: null });
        }
      }
      return { status: 200, data: { responses } };
    });
  }

  private handleRevoke(body: unknown): Promise<GraphResponse> {
    const token = body instanceof URLSearchParams ? (body.get('token') ?? 'unknown') : 'unknown';
    return this.dispatch('auth.revoke', token, () => ({ status: 200, data: {} }));
  }

  private handleSubsPatch(id: string, body: Record<string, unknown>): Promise<GraphResponse> {
    return this.dispatch('subs.patch', id, () => {
      const sub = this.subscriptions.get(id);
      if (!sub) return { status: 404, data: null };
      if (typeof body.expirationDateTime === 'string') {
        sub.expirationDateTime = body.expirationDateTime;
      }
      return { status: 200, data: sub };
    });
  }

  // ── axios wiring ───────────────────────────────────────────────────

  /** Route the jest-mocked axios methods into this fake Graph. */
  install(mockedAxios: AxiosMockLike): void {
    mockedAxios.isAxiosError.mockImplementation(
      (err: unknown): boolean =>
        typeof err === 'object' && err !== null && (err as { isAxiosError?: boolean }).isAxiosError === true,
    );

    mockedAxios.get.mockImplementation(async (url: string, config?: unknown) => {
      const params = (config as { params?: Record<string, unknown> } | undefined)?.params;
      const subMatch = /\/v1\.0\/subscriptions\/([^/?]+)/.exec(url);
      if (subMatch) return this.handleSubsGet(subMatch[1]);
      if (/\/v1\.0\/subscriptions\/?$/.test(url)) return this.handleSubsList();
      const userMatch = /\/v1\.0\/users\/([^/?]+)/.exec(url);
      if (userMatch) return this.handleUsersGet(userMatch[1]);
      if (/\/v1\.0\/users\/?$/.test(url)) return this.handleUsersLookup(params);
      throw new Error(`ChaosGraph: unhandled GET ${url}`);
    });

    mockedAxios.post.mockImplementation(async (url: string, body?: unknown) => {
      if (url.includes('/$batch')) return this.handleBatch(body as { requests?: BatchInnerRequest[] });
      if (/\/v1\.0\/subscriptions\/?$/.test(url)) return this.handleSubsCreate(body as Record<string, unknown>);
      if (url.includes('/logout')) return this.handleRevoke(body);
      throw new Error(`ChaosGraph: unhandled POST ${url}`);
    });

    mockedAxios.delete.mockImplementation(async (url: string) => {
      const subMatch = /\/v1\.0\/subscriptions\/([^/?]+)/.exec(url);
      if (subMatch) return this.handleSubsDelete(subMatch[1]);
      throw new Error(`ChaosGraph: unhandled DELETE ${url}`);
    });

    mockedAxios.patch.mockImplementation(async (url: string, body?: unknown) => {
      const subMatch = /\/v1\.0\/subscriptions\/([^/?]+)/.exec(url);
      if (subMatch) return this.handleSubsPatch(subMatch[1], (body ?? {}) as Record<string, unknown>);
      throw new Error(`ChaosGraph: unhandled PATCH ${url}`);
    });
  }
}
