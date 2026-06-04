/**
 * Minimal structural interface for an ioredis-compatible client.
 *
 * nestjs-outlook never imports `ioredis` directly. The host application
 * constructs the client and passes it in via `MicrosoftOutlookConfig.state.redis.client`.
 *
 * Any client satisfying this shape works: ioredis, ioredis-mock, a hand-rolled
 * test fake, or a redis-cluster wrapper.
 */
/**
 * Structural Redis interface. Methods are typed as `(...args: any[]) => Promise<any>`
 * because ioredis's real interface uses heavy overloads (callbacks, Buffer, RedisKey
 * brand types) that don't reduce to a single signature. We use `any` here only to
 * keep ioredis assignable to this port without forcing a hard dependency.
 */
/* eslint-disable @typescript-eslint/no-explicit-any -- ioredis's real signatures use heavy overloads (callbacks, Buffer, RedisKey brand types) that don't reduce to one signature; `any` keeps ioredis assignable to this structural port without a hard dependency. */
export interface RedisLike {
  ping(...args: any[]): Promise<any>;
  set(...args: any[]): Promise<any>;
  get(...args: any[]): Promise<any>;
  del(...args: any[]): Promise<any>;
  pexpire(...args: any[]): Promise<any>;
  eval(...args: any[]): Promise<any>;
  zadd(...args: any[]): Promise<any>;
  zremrangebyscore(...args: any[]): Promise<any>;
  zcard(...args: any[]): Promise<any>;
  hset(...args: any[]): Promise<any>;
  hgetall(...args: any[]): Promise<any>;
}
/* eslint-enable @typescript-eslint/no-explicit-any -- re-enable after the structural port above. */
