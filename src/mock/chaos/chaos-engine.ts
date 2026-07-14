import { SeededRandom } from './seeded-random';

/** Random disruption probabilities applied per Graph attempt (0 disables a category). */
export interface ChaosRates {
  /** 429 Too Many Requests with Retry-After header (retryable; costs ≥5s virtual backoff). */
  throttle429?: number;
  /** 503 Service Unavailable with Retry-After (retryable). */
  unavailable503?: number;
  /** 500 Internal Server Error (retryable via exponential backoff). */
  serverError500?: number;
  /** ECONNRESET-style network error (retryable). */
  networkError?: number;
}

/** A planned, deterministic failure targeting one logical key on one route. */
interface PlanEntry {
  status: number | 'network';
  /** 'always' keeps failing forever; a number fails that many times then succeeds. */
  remaining: number | 'always';
  /**
   * false (default): fail BEFORE the state mutation (at-most-once — request never reached Graph).
   * true: fail AFTER the mutation (at-least-once — Graph applied the change, the response was
   * lost). The second models the nastier real-world failure: a retried non-idempotent request.
   */
  afterExecute: boolean;
}

/** A chaos decision: the error to inject and whether it fires before or after the mutation. */
export interface ChaosDecision {
  error: ChaosHttpError;
  afterExecute: boolean;
}

/** An axios-shaped error the module's retry classifiers understand. */
export interface ChaosHttpError extends Error {
  isAxiosError: boolean;
  code?: string;
  response?: {
    status: number;
    headers: Record<string, string>;
    data: unknown;
  };
}

/** Build an axios-shaped HTTP error for a status (or a network error). */
export function buildChaosError(status: number | 'network'): ChaosHttpError {
  if (status === 'network') {
    const err = new Error('chaos: socket hang up') as ChaosHttpError;
    err.isAxiosError = true;
    err.code = 'ECONNRESET';
    return err;
  }
  const err = new Error(`chaos: HTTP ${status}`) as ChaosHttpError;
  err.isAxiosError = true;
  err.response = {
    status,
    // Retry-After present on throttle/unavailable; the module clamps it to >= 5s,
    // which shows up as *virtual* backoff time under fake timers.
    headers: status === 429 || status === 503 ? { 'retry-after': '0' } : {},
    data: { error: { code: `chaos_${status}`, message: `injected ${status}` } },
  };
  return err;
}

/**
 * Seeded chaos decision engine shared by the Graph and DB fakes.
 *
 * Two disruption modes compose:
 * - **Plans** (deterministic): `alwaysFail`/`failTimes` target a specific `(route, key)` so a
 *   test can compute exact expected outcomes (e.g. "exactly these 15 users fail").
 * - **Rates** (randomized but seeded): per-attempt probabilities of 429/503/500/network, for
 *   robustness runs where assertions are conservation laws and ceilings, not exact values.
 *
 * Plans are consulted before rates, so planned outcomes stay deterministic even in a storm.
 */
export class ChaosEngine {
  readonly random: SeededRandom;
  private readonly plans = new Map<string, PlanEntry>();

  constructor(
    seed: number,
    private rates: ChaosRates = {},
    private readonly latencyMs: { min: number; max: number } = { min: 0, max: 0 },
  ) {
    this.random = new SeededRandom(seed);
  }

  /** Change the random disruption rates mid-test (e.g. "the weather clears" for a rerun). */
  setRates(rates: ChaosRates): void {
    this.rates = rates;
  }

  private planKey(route: string, key: string): string {
    return `${route}|${key}`;
  }

  /** Fail every attempt for (route, key) with the given status — before the state mutation. */
  alwaysFail(route: string, key: string, status: number | 'network'): void {
    this.plans.set(this.planKey(route, key), { status, remaining: 'always', afterExecute: false });
  }

  /** Fail the first `times` attempts for (route, key), then let it succeed. */
  failTimes(route: string, key: string, times: number, status: number | 'network'): void {
    this.plans.set(this.planKey(route, key), { status, remaining: times, afterExecute: false });
  }

  /**
   * Fail the first `times` attempts AFTER the state mutation has been applied — the request
   * reached Graph and took effect, but the response was lost (at-least-once semantics). A
   * retried non-idempotent request (e.g. POST /subscriptions) then duplicates the mutation.
   */
  failTimesAfterExecute(route: string, key: string, times: number, status: number | 'network'): void {
    this.plans.set(this.planKey(route, key), { status, remaining: times, afterExecute: true });
  }

  /**
   * Decide the fate of one attempt. Returns the decision, or null to let it through untouched.
   * `plansOnly` skips the random rates — used by the DB fakes, where random disruption would
   * make planned outcomes nondeterministic (DB failures are always plan-targeted).
   *
   * A key whose `failTimes` plan is exhausted is FORCED to succeed (random rates are skipped),
   * so planned outcomes stay deterministic even inside a random storm.
   */
  decideFull(route: string, key: string, opts?: { plansOnly?: boolean }): ChaosDecision | null {
    const plan = this.plans.get(this.planKey(route, key));
    if (plan) {
      if (plan.remaining === 'always') {
        return { error: buildChaosError(plan.status), afterExecute: plan.afterExecute };
      }
      if (plan.remaining > 0) {
        plan.remaining -= 1;
        return { error: buildChaosError(plan.status), afterExecute: plan.afterExecute };
      }
      return null; // exhausted plan → forced success, immune to the random weather
    }
    if (opts?.plansOnly) return null;

    if (this.random.chance(this.rates.throttle429 ?? 0)) return { error: buildChaosError(429), afterExecute: false };
    if (this.random.chance(this.rates.unavailable503 ?? 0))
      return { error: buildChaosError(503), afterExecute: false };
    if (this.random.chance(this.rates.serverError500 ?? 0))
      return { error: buildChaosError(500), afterExecute: false };
    if (this.random.chance(this.rates.networkError ?? 0))
      return { error: buildChaosError('network'), afterExecute: false };
    return null;
  }

  /** Back-compat convenience: the injected error only (before-mutation), or null. */
  decide(route: string, key: string, opts?: { plansOnly?: boolean }): ChaosHttpError | null {
    return this.decideFull(route, key, opts)?.error ?? null;
  }

  /** Sampled latency for one call (milliseconds of *virtual* time under fake timers). */
  latency(): number {
    if (this.latencyMs.max <= 0) return 0;
    return this.random.int(this.latencyMs.min, this.latencyMs.max);
  }
}

/** setTimeout-based delay — under jest fake timers this costs virtual, not real, time. */
export function chaosDelay(ms: number): Promise<void> {
  if (ms <= 0) return Promise.resolve();
  return new Promise((resolve) => setTimeout(resolve, ms));
}
