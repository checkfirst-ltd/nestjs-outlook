# Running the demo over a public URL (zrok)

This demo talks to Microsoft Graph. Two flows need Microsoft to reach **your** app, not the
other way around:

- **OAuth callback** — after sign-in, Microsoft redirects the browser to
  `BACKEND_BASE_URL/auth/microsoft/callback`.
- **Webhooks** — calendar/email change notifications are POSTed by Graph to
  `BACKEND_BASE_URL/calendar/webhook` (and `/email/webhook`).

`http://localhost:8888` is not reachable from Microsoft, so for either flow you need a public
HTTPS URL. We use [zrok](https://zrok.io) (v2, the `zrok2` CLI) to get **one fixed URL** —
`https://nestjs-outlook-demo.shares.zrok.io` — that bridges to your local `:8888`. The URL stays
the same across runs, so you register it in Azure once.

> This mirrors the live-tier tunnel setup in `scheduleai/backend` (see its
> `src/test/calendar/README-live.md`), trimmed to what a runnable demo needs.

## Quick start

One-time setup (steps 1–3), then per-run (steps 4–6).

```bash
cd samples/nestjs-outlook-example

# 1. (once) enable zrok on this machine and reserve the fixed name
zrok2 enable <account-token>
zrok2 create name nestjs-outlook-demo        # → https://nestjs-outlook-demo.shares.zrok.io

# 2. (once) in the Azure app registration, add this redirect URI:
#    https://nestjs-outlook-demo.shares.zrok.io/auth/microsoft/callback
#    (append /<MICROSOFT_BASE_PATH> before /auth/... if you set that var)

# 3. (once) point the app at the public URL and let it start the tunnel itself
#    in .env:  BACKEND_BASE_URL=https://nestjs-outlook-demo.shares.zrok.io
#              ZROK_AUTOSTART=true

# 4. (per run) start the demo — it listens on :8888 AND opens the tunnel on boot
npm run start:dev

# 5. (per run, optional) confirm everything lines up
npm run zrok:preflight
```

Then open `https://nestjs-outlook-demo.shares.zrok.io` and run the OAuth login.

## Automatic tunnel (`ZROK_AUTOSTART`)

Set `ZROK_AUTOSTART=true` in `.env` and the app starts the tunnel itself right after it
begins listening (`src/main.ts` → `startZrokTunnel()` in `src/zrok-autostart.ts`) — no second
terminal, no `npm run zrok:share`. The tunnel is a child process of the app, so stopping the app
(Ctrl-C) tears the tunnel down with it and you don't leak a share.

- The reserved name defaults to the one in `BACKEND_BASE_URL` (e.g. `nestjs-outlook-demo` from
  `https://nestjs-outlook-demo.shares.zrok.io`); override with `ZROK_NAME`.
- Leave `ZROK_AUTOSTART` unset/`false` for plain localhost runs — the app skips the tunnel entirely.
- If `zrok2` isn't installed/found, the app logs a warning and keeps running on localhost.

`npm run zrok:share` still works if you'd rather run the tunnel in its own terminal.

## The npm scripts

| Script | What it does |
|--------|--------------|
| `npm run zrok:share` | Bridges the reserved name → local `:8888` (`zrok2 share public 8888 -n public:nestjs-outlook-demo --headless`). Override with `ZROK_PORT`, `ZROK_NAME`, `ZROK_NAMESPACE`. |
| `npm run zrok:preflight` | Read-only checklist: env present, `zrok2` installed, name reserved, app on `:8888`, and an end-to-end probe of the webhook endpoint through the public URL. Exits non-zero on a hard failure. |

Both resolve the `zrok2` binary from `$ZROK2_BIN`, then `zrok2` on `PATH`, then `~/bin/zrok2`.

## zrok v2 notes (`zrok2`)

zrok v2 **removed** the v1 `zrok reserve` / `share reserved` commands. The fixed URL now comes
from a reserved **name** in a **namespace**, bound to a public share with `-n <namespace>:<name>`:

```bash
zrok2 enable <account-token>                 # one-time per machine
zrok2 create name nestjs-outlook-demo        # reserves nestjs-outlook-demo.shares.zrok.io
zrok2 share public 8888 -n public:nestjs-outlook-demo --headless   # per run (this is what zrok:share runs)
```

- `zrok2 overview` lists reservations (look for `RESERVED=true`).
- The underlying share token is random each run, but the **public URL stays fixed** — that is
  what Azure needs.
- Tear down with `zrok2 delete name nestjs-outlook-demo` (and `zrok2 delete share <token>` for a
  stray share).
- `--headless` is required when the share runs in the background (no TTY).

## Port

The app listens on `:8888` (overridable via `PORT` — see `src/main.ts`). zrok must bridge to the
same port; `zrok:share` defaults to `8888` and reads `ZROK_PORT`/`PORT` if you change it. The
public webhook/redirect host stays the fixed reserved name regardless of the local port.

## Troubleshooting

Run `npm run zrok:preflight` first — it catches most of these.

| Symptom | Cause | Fix |
|---|---|---|
| preflight `[3]` "not reserved" | the fixed name doesn't exist yet | `zrok2 create name nestjs-outlook-demo` |
| preflight `[5]` 404 "no share running" | name resolves but no share is bridged | `npm run zrok:share` |
| preflight `[5]` 502 / 503 | share up, app not listening | `npm run start:dev` |
| OAuth ends on a Microsoft error page | redirect URI not registered, or `BACKEND_BASE_URL` ≠ the zrok URL | register `…/auth/microsoft/callback` in Azure; set `BACKEND_BASE_URL` to the zrok URL |
| webhook subscription fails to create | Graph can't validate the notification URL (tunnel down) | start `zrok:share` before triggering the subscription |
| `zrok2 ... already in use` | a previous share/binding is still registered | `zrok2 delete share <token>` (find it via `zrok2 overview`), then re-share |
| `open /dev/tty: device not configured` | `zrok2 share` TUI needs a terminal | already handled — `zrok:share` passes `--headless` |
