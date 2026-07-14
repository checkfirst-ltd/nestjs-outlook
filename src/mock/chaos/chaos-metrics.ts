/**
 * Consumption/behaviour metrics collected by the chaos fakes.
 *
 * Everything is count- or ordering-based (robust in CI); wall-clock numbers are reported for
 * visibility but never asserted. "Virtual" time is jest fake-timer time — what the flow *would*
 * have waited on latency and retry backoff in production.
 */
export class ChaosMetrics {
  /** Attempts per route (each retry counts — attempts − logical ops = retries). */
  readonly attempts = new Map<string, number>();
  /** Injected errors per `route|status`. */
  readonly injected = new Map<string, number>();
  /** Attempts per `route|key` — lets tests assert exact retry counts for planned keys. */
  readonly perKeyAttempts = new Map<string, number>();
  /** DB fake calls per method name. */
  readonly dbCalls = new Map<string, number>();
  /** Ordered trace of notable operations (graph routes, db writes, lifecycle steps). */
  readonly timeline: string[] = [];

  private inFlight = 0;
  peakInFlight = 0;

  private readonly startedRealMs = performance.now();
  private readonly startedVirtualMs = Date.now();
  private readonly startedHeap = process.memoryUsage().heapUsed;

  private bump(map: Map<string, number>, key: string): void {
    map.set(key, (map.get(key) ?? 0) + 1);
  }

  enter(route: string, key: string): void {
    this.inFlight += 1;
    this.peakInFlight = Math.max(this.peakInFlight, this.inFlight);
    this.bump(this.attempts, route);
    this.bump(this.perKeyAttempts, `${route}|${key}`);
    this.timeline.push(`graph:${route}`);
  }

  exit(): void {
    this.inFlight -= 1;
  }

  recordInjected(route: string, status: number | 'network'): void {
    this.bump(this.injected, `${route}|${status}`);
  }

  recordDb(method: string): void {
    this.bump(this.dbCalls, `db.${method}`);
    this.timeline.push(`db:${method}`);
  }

  mark(step: string): void {
    this.timeline.push(step);
  }

  attemptsFor(route: string): number {
    return this.attempts.get(route) ?? 0;
  }

  attemptsForKey(route: string, key: string): number {
    return this.perKeyAttempts.get(`${route}|${key}`) ?? 0;
  }

  injectedFor(route: string, status?: number | 'network'): number {
    if (status !== undefined) return this.injected.get(`${route}|${status}`) ?? 0;
    let total = 0;
    for (const [key, count] of this.injected) {
      if (key.startsWith(`${route}|`)) total += count;
    }
    return total;
  }

  totalInjected(): number {
    let total = 0;
    for (const count of this.injected.values()) total += count;
    return total;
  }

  dbCallsFor(method: string): number {
    return this.dbCalls.get(`db.${method}`) ?? 0;
  }

  /** Index of the last timeline entry matching a prefix (-1 when absent). */
  lastIndexOf(prefix: string): number {
    for (let i = this.timeline.length - 1; i >= 0; i--) {
      if (this.timeline[i].startsWith(prefix)) return i;
    }
    return -1;
  }

  /** Human-readable consumption summary (report-only; never asserted). */
  report(label: string): string {
    const realMs = Math.round(performance.now() - this.startedRealMs);
    const virtualMs = Date.now() - this.startedVirtualMs;
    const heapMb = (process.memoryUsage().heapUsed - this.startedHeap) / (1024 * 1024);

    const lines: string[] = [
      `── chaos report: ${label} ──`,
      `real cpu time      : ${realMs}ms`,
      `virtual wall-clock : ${virtualMs}ms (latency + retry backoff that production would wait)`,
      `heap delta         : ${heapMb.toFixed(1)}MB`,
      `peak graph in-flight: ${this.peakInFlight}`,
      `graph attempts     : ${[...this.attempts.entries()].map(([r, n]) => `${r}=${n}`).join(', ') || 'none'}`,
      `injected errors    : ${[...this.injected.entries()].map(([r, n]) => `${r}=${n}`).join(', ') || 'none'}`,
      `db calls           : ${[...this.dbCalls.entries()].map(([r, n]) => `${r}=${n}`).join(', ') || 'none'}`,
    ];
    return lines.join('\n');
  }
}
