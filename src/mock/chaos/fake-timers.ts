/**
 * Fake-timer driver for chaos flows.
 *
 * Every sleep in the module funnels through `setTimeout` (chaos latency, `retry.util`'s
 * exponential backoff, the ≥5s Retry-After waits on 429/503). Faking timers makes those waits
 * cost *virtual* time only — a chaos storm that would take minutes of real backoff completes in
 * seconds of CPU — while `Date.now()` still advances so staleness/expiry math stays coherent.
 */

/** Install fake timers (call before building a chaos world so the clock is coherent). */
export function installChaosTimers(): void {
  jest.useFakeTimers({
    // Keep the real microtask/immediate machinery and real perf counters: the drain loop
    // below uses setImmediate hops, and ChaosMetrics measures real CPU via performance.now.
    doNotFake: ['nextTick', 'setImmediate', 'queueMicrotask', 'performance', 'hrtime'],
  });
}

export function uninstallChaosTimers(): void {
  jest.useRealTimers();
}

const MAX_TICKS = 200_000;
const MAX_IDLE_HOPS = 50;

/**
 * Pump fake timers until `promise` settles, then return/throw its result.
 * Throws if the flow stalls (pending forever with no timers) — a deadlock detector.
 */
export async function drain<T>(promise: Promise<T>): Promise<T> {
  let settled = false;
  const guarded = promise.then(
    (value) => {
      settled = true;
      return value;
    },
    (error: unknown) => {
      settled = true;
      throw error;
    },
  );
  // Avoid unhandled-rejection noise while we pump timers; the caller awaits `guarded`.
  guarded.catch(() => undefined);

  let idleHops = 0;
  for (let tick = 0; tick < MAX_TICKS && !settled; tick++) {
    if (jest.getTimerCount() > 0) {
      idleHops = 0;
      await jest.runOnlyPendingTimersAsync();
    } else {
      // No timers pending — give promise chains a real macrotask hop to progress.
      await new Promise<void>((resolve) => setImmediate(resolve));
      idleHops += 1;
      if (idleHops > MAX_IDLE_HOPS && !settled && jest.getTimerCount() === 0) {
        throw new Error('drain(): flow stalled — promise pending with no fake timers scheduled');
      }
    }
  }
  if (!settled) {
    throw new Error(`drain(): flow did not settle within ${MAX_TICKS} timer ticks`);
  }
  return guarded;
}

/**
 * Pump fake timers until `predicate()` turns true — for detached background flows
 * (the 202 endpoints) where there is no promise to await, only an observable effect
 * (an emitted event, a timeline mark).
 */
export async function drainUntil(predicate: () => boolean): Promise<void> {
  let idleHops = 0;
  for (let tick = 0; tick < MAX_TICKS; tick++) {
    if (predicate()) return;
    if (jest.getTimerCount() > 0) {
      idleHops = 0;
      await jest.runOnlyPendingTimersAsync();
    } else {
      await new Promise<void>((resolve) => setImmediate(resolve));
      idleHops += 1;
      if (idleHops > MAX_IDLE_HOPS && !predicate() && jest.getTimerCount() === 0) {
        throw new Error('drainUntil(): condition never became true and no fake timers are scheduled');
      }
    }
  }
  throw new Error(`drainUntil(): condition not reached within ${MAX_TICKS} timer ticks`);
}
