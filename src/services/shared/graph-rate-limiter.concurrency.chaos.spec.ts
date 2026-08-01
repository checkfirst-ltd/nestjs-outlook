import { Logger } from '@nestjs/common';
import axios from 'axios';
import {
  ChaosEngine,
  ChaosGraph,
  ChaosMetrics,
  AxiosMockLike,
  installChaosTimers,
  uninstallChaosTimers,
  drain,
} from '../../mock/chaos';
import { GraphRateLimiterService } from './graph-rate-limiter.service';
import { InMemoryOutlookRateLimitStore } from './outlook-rate-limit.store';
import { executeGraphApiCall } from '../../utils/outlook-api-executor.util';

jest.mock('axios');
const mockedAxios = axios as jest.Mocked<typeof axios> as unknown as AxiosMockLike;

jest.setTimeout(120_000);

/**
 * Reproduces the production incident: a tenant-wide connect-all fans out many concurrent Graph
 * reads onto the SAME mailbox (webhook-triggered delta read + initial sync + subscription/health
 * reads), overrunning Outlook's per-(app, mailbox) concurrency ceiling of 4 → HTTP 429
 * "Application is over its MailboxConcurrency limit."
 *
 * The wiring under test is the real GraphRateLimiterService driving executeGraphApiCall — the
 * choke point every Graph call flows through. The existing limiter caps *rate* (4 req/sec/user)
 * but not *concurrency*: when a Graph call outlives the 1-second window, in-flight requests on a
 * mailbox pile past 4 and the mailbox throttles. The fix (A1) adds a per-mailbox in-flight
 * semaphore so app-side concurrency never reaches Outlook's ceiling.
 */
describe('GraphRateLimiterService — per-mailbox concurrency (MailboxConcurrency 429)', () => {
  const SEED = Number(process.env.CHAOS_SEED ?? 20260731);
  const MAILBOX = 'ms-connect-all-victim';
  const MAILBOX_CONCURRENCY_LIMIT = 4; // Outlook's real per-(app, mailbox) ceiling.
  const CONCURRENT_READS = 12; // one tenant's auditors all landing on the same mailbox at once.

  let graph: ChaosGraph;
  let metrics: ChaosMetrics;
  let limiter: GraphRateLimiterService;

  beforeAll(() => {
    Logger.overrideLogger(false);
  });

  beforeEach(() => {
    jest.clearAllMocks();
    installChaosTimers();

    metrics = new ChaosMetrics();
    // Latency deliberately exceeds the limiter's 1-second rate window so that rate-limiting alone
    // cannot bound concurrency — reads outlive the window and pile up on the mailbox.
    const engine = new ChaosEngine(SEED, {}, { min: 1500, max: 2500 });
    graph = new ChaosGraph(engine, metrics);
    graph.mailboxConcurrencyLimit = MAILBOX_CONCURRENCY_LIMIT;
    graph.install(mockedAxios);

    limiter = new GraphRateLimiterService(new InMemoryOutlookRateLimitStore());
  });

  afterEach(() => {
    uninstallChaosTimers();
  });

  /** One calendar read against the victim mailbox, throttled through the real limiter + executor. */
  const readMailbox = () =>
    executeGraphApiCall(
      () => axios.get(`https://graph.microsoft.com/v1.0/users/${MAILBOX}/events/delta`),
      {
        rateLimiter: limiter,
        userId: MAILBOX,
        resourceName: `users/${MAILBOX}/events/delta`,
        logger: { warn: () => undefined, error: () => undefined },
      },
    );

  it('bounds concurrent reads on a single mailbox below Outlook\'s ceiling (no MailboxConcurrency 429)', async () => {
    const settled = await drain(
      Promise.allSettled(Array.from({ length: CONCURRENT_READS }, () => readMailbox())),
    );

    const succeeded = settled.filter((r) => r.status === 'fulfilled').length;

    // The app must never push a mailbox past Outlook's concurrency ceiling: zero throttle
    // rejections, and observed in-flight per mailbox stays under the limit.
    expect(metrics.mailboxConcurrencyRejections).toBe(0);
    expect(metrics.injectedFor('mailbox.read', 429)).toBe(0);
    expect(metrics.peakMailboxInFlight()).toBeLessThan(MAILBOX_CONCURRENCY_LIMIT);

    // And every read still completes (work is throttled, not dropped).
    expect(succeeded).toBe(CONCURRENT_READS);

    console.log(metrics.report(`mailbox-concurrency N=${CONCURRENT_READS} seed=${SEED}`));
  });
});
