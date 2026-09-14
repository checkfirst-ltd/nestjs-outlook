import { extractErrorInfo, retryWithBackoff } from './retry.util';

/** Graph SDK shaped error — detail lives on statusCode/code/body, not on .message. */
const graphError = (statusCode: number, code = 'unknownError') =>
  Object.assign(new Error(''), { statusCode, code, body: 'throttled' });

describe('extractErrorInfo', () => {
  it('reads a Graph SDK error from statusCode/code/body rather than .message', () => {
    expect(extractErrorInfo(graphError(429, 'ApplicationThrottled'))).toEqual({
      statusCode: 429,
      code: 'ApplicationThrottled',
      message: 'throttled',
      type: 'graph_api_error',
    });
  });

  it('classifies statusCode -1 as a network error', () => {
    expect(extractErrorInfo(graphError(-1)).type).toBe('network_error');
  });

  it('reads an axios-shaped error', () => {
    const err = {
      response: { status: 503, data: { error: { code: 'ServiceUnavailable', message: 'try later' } } },
    };

    expect(extractErrorInfo(err)).toEqual({
      statusCode: 503,
      code: 'ServiceUnavailable',
      message: 'try later',
      type: 'axios_api_error',
    });
  });
});

describe('retryWithBackoff', () => {
  const warn = jest.fn();
  const logger = { warn };

  beforeEach(() => {
    jest.useFakeTimers();
    warn.mockClear();
  });

  afterEach(() => {
    jest.useRealTimers();
  });

  /** Drive an in-flight retryWithBackoff past all of its pending sleeps. */
  const settle = async (promise: Promise<unknown>) => {
    const caught = promise.catch((e: unknown) => e);
    for (let i = 0; i < 20; i++) {
      await Promise.resolve();
      jest.runAllTimers();
    }
    return caught;
  };

  it('makes exactly maxRetries+1 attempts before giving up', async () => {
    const operation = jest.fn().mockRejectedValue(graphError(429));

    await settle(retryWithBackoff(operation, { maxRetries: 2, logger }));

    expect(operation).toHaveBeenCalledTimes(3);
  });

  it('stops immediately on a non-retryable status', async () => {
    const operation = jest.fn().mockRejectedValue(graphError(403));

    await settle(retryWithBackoff(operation, { maxRetries: 10, logger }));

    expect(operation).toHaveBeenCalledTimes(1);
    expect(warn.mock.calls[0][0]).toContain('status=403');
  });

  it('puts the status and code in the message, not only the context object', async () => {
    const operation = jest.fn().mockRejectedValue(graphError(429, 'ApplicationThrottled'));

    await settle(retryWithBackoff(operation, { maxRetries: 1, logger, operationName: 'series fetch' }));

    const messages = warn.mock.calls.map((c: unknown[]) => String(c[0]));

    // A Nest Logger takes a string context as its second argument and drops an
    // object, so anything only present there never reaches the log line.
    expect(messages.some((m) => m.includes('status=429') && m.includes('ApplicationThrottled'))).toBe(true);
    expect(messages.some((m) => m.includes('Max retries (1) exceeded for series fetch'))).toBe(true);
  });

  it('returns the value as soon as an attempt succeeds', async () => {
    const operation = jest
      .fn()
      .mockRejectedValueOnce(graphError(503))
      .mockResolvedValue('ok');

    await expect(settle(retryWithBackoff(operation, { maxRetries: 5, logger }))).resolves.toBe('ok');
    expect(operation).toHaveBeenCalledTimes(2);
  });
});
