import { spawn, execFileSync, ChildProcess } from 'child_process';
import { homedir } from 'os';
import { join } from 'path';

/**
 * Optionally start a zrok v2 tunnel that bridges the fixed public URL -> this app's
 * local port, so Microsoft Graph can reach the OAuth callback and webhooks without a
 * second terminal running `npm run zrok:share`.
 *
 * Opt-in via ZROK_AUTOSTART=true so plain localhost runs are unaffected. Everything
 * else (reserved name, Azure redirect URI, BACKEND_BASE_URL) is still one-time setup —
 * see ZROK.md. This only removes the per-run "start the tunnel by hand" step.
 *
 * Overrides: ZROK_NAME, ZROK_NAMESPACE, ZROK2_BIN. When ZROK_NAME is unset the reserved
 * name is derived from BACKEND_BASE_URL if it is a *.shares.zrok.io URL.
 */
export function startZrokTunnel(port: number | string): ChildProcess | undefined {
  if (process.env.ZROK_AUTOSTART !== 'true') return undefined;

  const name = resolveReservedName();
  const namespace = process.env.ZROK_NAMESPACE || 'public';
  const zrokBin = resolveZrokBin();

  if (!zrokBin) {
    console.warn(
      '⚠️  ZROK_AUTOSTART=true but the zrok2 CLI was not found. Install zrok v2 and run ' +
        '`zrok2 enable <token>` (or set $ZROK2_BIN). Skipping tunnel — the app still runs on localhost.',
    );
    return undefined;
  }

  console.log(
    `▶ zrok: bridging https://${name}.shares.zrok.io  ->  http://localhost:${port}`,
  );

  const child = spawn(
    zrokBin,
    ['share', 'public', String(port), '-n', `${namespace}:${name}`, '--headless'],
    { stdio: 'inherit' },
  );

  child.on('error', (err) => {
    console.warn(`⚠️  zrok tunnel failed to start: ${err.message}. Local app still runs.`);
  });
  child.on('exit', (code) => {
    if (code && code !== 0) {
      console.warn(
        `⚠️  zrok tunnel exited with code ${code}. The public URL is down; the local app keeps running. ` +
          "A stray share may still be registered — see ZROK.md ('already in use').",
      );
    }
  });

  // Tear the tunnel down whenever this process goes away, so we don't leak shares.
  const stop = () => {
    if (!child.killed) child.kill();
  };
  process.on('exit', stop);
  process.on('SIGINT', stop);
  process.on('SIGTERM', stop);

  return child;
}

/** ZROK_NAME wins; otherwise pull the name out of a *.shares.zrok.io BACKEND_BASE_URL. */
function resolveReservedName(): string {
  if (process.env.ZROK_NAME) return process.env.ZROK_NAME;
  const base = process.env.BACKEND_BASE_URL;
  if (base) {
    try {
      const host = new URL(base).hostname; // nestjs-outlook-demo.shares.zrok.io
      if (host.endsWith('.zrok.io')) return host.split('.')[0];
    } catch {
      /* fall through to default */
    }
  }
  return 'nestjs-outlook-demo';
}

/** Resolve the zrok v2 binary the same way scripts/zrok-share.sh and the preflight do. */
function resolveZrokBin(): string | null {
  const candidates = [
    process.env.ZROK2_BIN,
    'zrok2',
    join(homedir(), 'bin', 'zrok2'),
  ].filter(Boolean) as string[];
  for (const bin of candidates) {
    try {
      execFileSync(bin, ['version'], { stdio: 'ignore' });
      return bin;
    } catch {
      /* keep looking */
    }
  }
  return null;
}
