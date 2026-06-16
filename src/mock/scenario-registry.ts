import { MockScenario } from './interfaces';

export class ScenarioRegistry {
  private static scenarios = new Map<string, MockScenario>();

  static register(scenario: MockScenario): void {
    ScenarioRegistry.scenarios.set(scenario.name, scenario);
  }

  static get(name: string): MockScenario | undefined {
    return ScenarioRegistry.scenarios.get(name);
  }

  static list(): string[] {
    return Array.from(ScenarioRegistry.scenarios.keys());
  }

  static has(name: string): boolean {
    return ScenarioRegistry.scenarios.has(name);
  }
}
