/**
 * Optional instrumentation interface for recording custom events,
 * metrics, and enriching log context from the host application.
 *
 * Host apps provide an implementation via the OUTLOOK_INSTRUMENTATION
 * DI token. When not provided, nestjs-outlook operates without instrumentation.
 */
export interface OutlookInstrumentation {
  // Custom event recording (e.g., New Relic recordCustomEvent)
  recordCustomEvent(
    eventType: string,
    attributes: Record<string, string | number | boolean>,
  ): void;

  // Metric recording (e.g., New Relic recordMetric)
  recordMetric(name: string, value: number): void;

  // Error recording (e.g., New Relic noticeError)
  noticeError(
    error: Error,
    attributes?: Record<string, string | number | boolean>,
  ): void;

  // Transaction-level custom attributes (e.g., New Relic addCustomAttributes)
  addCustomAttributes(
    params: Record<string, string | number | boolean>,
  ): void;

  // Log context enrichment — sets values in the host app's async context
  // so all subsequent Logger calls include these fields automatically
  setLogContext(key: string, value: string | number | boolean): void;
}

export const OUTLOOK_INSTRUMENTATION = 'OUTLOOK_INSTRUMENTATION';
