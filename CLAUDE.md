# CLAUDE.md

Guidance for Claude Code when working in the `@checkfirst/nestjs-outlook` repository.

## What this is

An opinionated NestJS module wrapping the Microsoft Graph API — OAuth authentication, calendar,
email, and webhook subscriptions. Published to npm as `@checkfirst/nestjs-outlook`.

## Documentation

Full documentation lives in [`docs/`](docs/) and follows the **Documentation Engineering
Protocol (DEP)**. A [`.docspec`](.docspec) file is present, so the DEP protocol is active —
query, navigate, and validate docs with the `dep` CLI rather than editing frontmatter by hand.

- **Root hub:** [`docs/index.md`](docs/index.md) — grouped by audience and type
- **Onboarding:** [`docs/tutorials/getting-started.md`](docs/tutorials/getting-started.md)
- **Task guides:** [`docs/how-to/`](docs/how-to/)
- **API reference:** [`docs/reference/api-overview.md`](docs/reference/api-overview.md)
- **Internals & rationale:** [`docs/explanation/`](docs/explanation/), [`docs/decision-records/`](docs/decision-records/)

Keep `docs/` in sync with code: each document's `depends_on` lists the source files that, when
changed, should trigger a re-verify. After editing docs:

```bash
dep validate --root .   # check metadata, type purity, graph integrity
dep index --root .      # regenerate index files from metadata
```

> Note: re-running `dep index` resets the root `docs/index.md` owner to `@dep-core`. Restore it
> with `dep set docs/index.md --owner "@checkfirst-ltd"`.

## Common commands

```bash
npm run build        # rimraf dist && tsc
npm test             # jest (npm run test:watch, npm run test:cov)
npm run lint         # eslint (npm run lint:fix to autofix)
npm run format       # prettier --write "src/**/*.ts"
npm run dev          # watch build + yalc push (local linking into a host app)
```

## Source layout (`src/`)

- `services/` — `auth`, `calendar`, `email`, `subscription`, and `shared` (locking, rate
  limiting, delta sync, user-id conversion)
- `controllers/` — OAuth callback + calendar/email webhook endpoints
- `entities/`, `migrations/`, `repositories/` — TypeORM persistence (run migrations before use)
- `enums/`, `interfaces/`, `guards/` — public contracts and the webhook `clientState` guard
- `index.ts` — the public export surface

## Conventions

- Conventional commits; release-please drives `CHANGELOG.md` and versioning.
- This package uses **npm** (not bun) — match the existing tooling.
