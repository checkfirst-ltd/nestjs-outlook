/**
 * Map over `items` running at most `concurrency` `worker` calls at a time.
 *
 * A fixed pool of worker loops each pull the next index from a shared cursor until
 * the list is exhausted — bounded concurrency without recursion and without
 * allocating one promise per item up front. Use it to speed up I/O-bound fan-out
 * (e.g. per-user Microsoft Graph calls) while keeping a ceiling that respects Graph
 * rate limits.
 *
 * Results preserve input order. A rejecting `worker` rejects the whole call, so wrap
 * per-item work in try/catch when partial failures should be tolerated.
 *
 * @param items - Items to process.
 * @param concurrency - Maximum number of concurrent `worker` calls (clamped to >= 1).
 * @param worker - Async function invoked with each item and its index.
 * @returns Results in the same order as `items`.
 */
export async function mapWithConcurrency<T, R>(
  items: readonly T[],
  concurrency: number,
  worker: (item: T, index: number) => Promise<R>,
): Promise<R[]> {
  const results: R[] = new Array<R>(items.length);
  if (items.length === 0) {
    return results;
  }

  // No point spinning up more runners than there is work.
  const poolSize = Math.max(1, Math.min(Math.floor(concurrency), items.length));
  let cursor = 0;

  const runner = async (): Promise<void> => {
    // Each runner keeps claiming the next unprocessed index until none remain.
    // `cursor++` is safe here because JS runs this synchronously between awaits.
    let index = cursor++;
    while (index < items.length) {
      results[index] = await worker(items[index], index);
      index = cursor++;
    }
  };

  const runners: Promise<void>[] = [];
  for (let i = 0; i < poolSize; i++) {
    runners.push(runner());
  }
  await Promise.all(runners);

  return results;
}
