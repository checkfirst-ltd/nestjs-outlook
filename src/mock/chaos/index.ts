/**
 * Chaos test harness (test-only; `src/mock` is excluded from the published build).
 *
 * Wires the REAL provisioning / subscription / tenant-user / health services to an in-memory
 * Microsoft Graph and database behind a seeded chaos layer — latency, throttling, server
 * errors, network drops, and deterministic per-key fail-plans — so host-facing flows can be
 * exercised at scale and their behaviour, consumption, and recovery verified.
 */
export * from './seeded-random';
export * from './chaos-engine';
export * from './chaos-metrics';
export * from './chaos-graph';
export * from './chaos-db';
export * from './chaos-world';
export * from './fake-timers';
