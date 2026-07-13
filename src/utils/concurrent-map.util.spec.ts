import { mapWithConcurrency } from './concurrent-map.util';

describe('mapWithConcurrency', () => {
  it('returns an empty array for empty input without invoking the worker', async () => {
    const worker = jest.fn();
    const result = await mapWithConcurrency([], 4, worker);
    expect(result).toEqual([]);
    expect(worker).not.toHaveBeenCalled();
  });

  it('maps every item and preserves input order', async () => {
    const items = [1, 2, 3, 4, 5];
    const result = await mapWithConcurrency(items, 2, async (n) => n * 10);
    expect(result).toEqual([10, 20, 30, 40, 50]);
  });

  it('passes the index to the worker', async () => {
    const result = await mapWithConcurrency(['a', 'b', 'c'], 3, async (item, index) => `${index}:${item}`);
    expect(result).toEqual(['0:a', '1:b', '2:c']);
  });

  it('never exceeds the concurrency ceiling', async () => {
    let inFlight = 0;
    let maxInFlight = 0;
    const items = Array.from({ length: 12 }, (_, i) => i);

    await mapWithConcurrency(items, 3, async (n) => {
      inFlight++;
      maxInFlight = Math.max(maxInFlight, inFlight);
      // Yield so overlapping workers are observable.
      await new Promise((resolve) => setTimeout(resolve, 1));
      inFlight--;
      return n;
    });

    expect(maxInFlight).toBeLessThanOrEqual(3);
    expect(maxInFlight).toBeGreaterThan(1);
  });

  it('clamps concurrency to at least 1', async () => {
    const result = await mapWithConcurrency([1, 2, 3], 0, async (n) => n);
    expect(result).toEqual([1, 2, 3]);
  });

  it('rejects if a worker rejects', async () => {
    await expect(
      mapWithConcurrency([1, 2, 3], 2, async (n) => {
        if (n === 2) {
          throw new Error('boom');
        }
        return n;
      }),
    ).rejects.toThrow('boom');
  });
});
