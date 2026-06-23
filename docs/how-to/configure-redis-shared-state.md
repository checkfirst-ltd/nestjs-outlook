---
dep:
  type: how-to
  audience: [app-developer, library-integrator]
  owner: "@checkfirst-ltd"
  created: 2026-06-23
  last_verified: 2026-06-23T12:00:00+03:00
  confidence: high
  depends_on:
    - src/interfaces/config/outlook-config.interface.ts
    - src/microsoft-outlook.module.ts
    - src/interfaces/state/redis-like.interface.ts
  tags: [redis, state, multi-container, configuration]
  links:
    - target: ../reference/configuration.md
      rel: USES
    - target: ../reference/shared-state-stores.md
      rel: USES
    - target: ../explanation/shared-state-and-concurrency.md
      rel: NEXT
    - target: ../tutorials/getting-started.md
      rel: REQUIRES
---

# Configure Redis Shared State

**Goal:** Make locking and rate-limit budgets consistent across multiple containers by backing
the module's shared state with Redis.

Use this when you run more than one instance of your app. With the default in-memory backend,
each container keeps its own locks and counters, which reintroduces cross-container concurrency
problems.

## Steps

### 1. Provide a Redis client

The module never imports `ioredis`. Construct a client in your app and pass it in. Any client
matching the `RedisLike` shape works.

```typescript
import IORedis from 'ioredis';

const redisClient = new IORedis(process.env.REDIS_URL);
```

### 2. Pass the client to the module

```typescript
MicrosoftOutlookModule.forRoot({
  clientId: process.env.MS_CLIENT_ID,
  clientSecret: process.env.MS_CLIENT_SECRET,
  redirectPath: 'auth/microsoft/callback',
  backendBaseUrl: 'https://your-api.example.com',
  basePath: 'api/v1',
  state: {
    redis: {
      client: redisClient,
      keyPrefix: 'outlook:',
      required: true,
    },
  },
});
```

### 3. Decide the failure policy

- `required: false` (default) — if the Redis `PING` probe fails at startup, the module logs a
  warning and falls back to in-memory.
- `required: true` — a failed probe throws during module init, so your orchestrator restarts
  the container instead of running in a degraded mode. Recommended for production.

### 4. Use `forRootAsync` if the client comes from DI

```typescript
MicrosoftOutlookModule.forRootAsync({
  inject: [REDIS_CLIENT],
  useFactory: (client) => ({
    clientId: process.env.MS_CLIENT_ID,
    clientSecret: process.env.MS_CLIENT_SECRET,
    redirectPath: 'auth/microsoft/callback',
    backendBaseUrl: 'https://your-api.example.com',
    state: { redis: { client, required: true } },
  }),
});
```

## Verify

- On startup, the logs read `OutlookLockStore backend: redis` and `OutlookRateLimitStore backend: redis`.
- With `required: true`, stopping Redis and restarting the app causes module init to fail loudly.
- Redis shows keys under your prefix (e.g. `outlook:lock:*`, `outlook:rl:*`).

## Related

- [Configuration reference](../reference/configuration.md) — the `state.redis` fields.
- [Shared-state stores reference](../reference/shared-state-stores.md) — the store interfaces.
- [Shared state and concurrency](../explanation/shared-state-and-concurrency.md) — why this matters.
