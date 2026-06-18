import { ScenarioRegistry } from '../scenario-registry';
import { happyPathScenario } from './happy-path.scenario';
import { tokenExpiredScenario } from './token-expired.scenario';
import { rateLimitedScenario } from './rate-limited.scenario';
import { mailboxInactiveScenario } from './mailbox-inactive.scenario';
import { emptyCalendarScenario } from './empty-calendar.scenario';
import { batchPartialFailureScenario } from './batch-partial-failure.scenario';

// Auto-register all built-in scenarios
ScenarioRegistry.register(happyPathScenario);
ScenarioRegistry.register(tokenExpiredScenario);
ScenarioRegistry.register(rateLimitedScenario);
ScenarioRegistry.register(mailboxInactiveScenario);
ScenarioRegistry.register(emptyCalendarScenario);
ScenarioRegistry.register(batchPartialFailureScenario);

export {
  happyPathScenario,
  tokenExpiredScenario,
  rateLimitedScenario,
  mailboxInactiveScenario,
  emptyCalendarScenario,
  batchPartialFailureScenario,
};
