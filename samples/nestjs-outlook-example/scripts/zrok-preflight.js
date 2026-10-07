/* eslint-disable */
/**
 * Preflight for running this demo over a zrok tunnel. Tells you exactly what is
 * missing before you start a real OAuth / webhook round-trip, so you don't
 * discover it mid-demo. Read-only — it never changes state.
 *
 * Run:  npm run zrok:preflight
 *
 * Prints a ✅/⚠️/❌ checklist and exits non-zero if any REQUIRED check fails.
 * See ZROK.md for the full setup.
 */
const os = require('os');
const path = require('path');
const net = require('net');
const { execFileSync } = require('child_process');

require('dotenv').config();

const DEFAULT_PORT = Number(process.env.PORT || '8888');

let hardFail = false;
const ok = (m) => console.log(`  ✅ ${m}`);
const warn = (m) => console.log(`  ⚠️  ${m}`);
const fail = (m) => {
  console.log(`  ❌ ${m}`);
  hardFail = true;
};

function resolveZrokBin() {
  const candidates = [
    process.env.ZROK2_BIN,
    'zrok2',
    path.join(os.homedir(), 'bin', 'zrok2'),
  ].filter(Boolean);
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

function reservedNameFromBaseUrl(baseUrl) {
  try {
    return new URL(baseUrl).hostname.split('.')[0]; // nestjs-outlook-demo.shares.zrok.io → nestjs-outlook-demo
  } catch {
    return null;
  }
}

// The library builds its webhook/callback URLs as BACKEND_BASE_URL + /<basePath>/...
// Mirror that so the probe hits the exact path Microsoft will use.
function webhookUrl(baseUrl) {
  const base = baseUrl.replace(/\/+$/, '');
  const basePath = (process.env.MICROSOFT_BASE_PATH || '').replace(/^\/+|\/+$/g, '');
  const prefix = basePath ? `/${basePath}` : '';
  return `${base}${prefix}/calendar/webhook?validationToken=ping`;
}

function portInUse(port) {
  return new Promise((resolve) => {
    const server = net.createServer();
    server.once('error', () => resolve(true));
    server.once('listening', () => server.close(() => resolve(false)));
    server.listen(port, '0.0.0.0');
  });
}

async function main() {
  console.log('\nzrok demo — preflight\n=====================');

  const baseUrl = process.env.BACKEND_BASE_URL;
  const clientId = process.env.MICROSOFT_CLIENT_ID;
  const clientSecret = process.env.MICROSOFT_CLIENT_SECRET;

  // 1. Env present
  console.log('\n[1] Environment (.env)');
  if (baseUrl && /^https:\/\/.+\.zrok\.io/.test(baseUrl)) {
    ok(`BACKEND_BASE_URL = ${baseUrl}`);
  } else if (baseUrl && /^https:\/\//.test(baseUrl)) {
    warn(`BACKEND_BASE_URL = ${baseUrl} (https but not a *.zrok.io URL — fine if it's a custom domain)`);
  } else {
    fail('BACKEND_BASE_URL must be your fixed public https URL (e.g. https://nestjs-outlook-demo.shares.zrok.io)');
  }
  clientId ? ok('MICROSOFT_CLIENT_ID present') : fail('MICROSOFT_CLIENT_ID missing');
  clientSecret ? ok('MICROSOFT_CLIENT_SECRET present') : fail('MICROSOFT_CLIENT_SECRET missing');

  const basePath = (process.env.MICROSOFT_BASE_PATH || '').replace(/^\/+|\/+$/g, '');
  const redirectPath = (process.env.MICROSOFT_REDIRECT_PATH || 'auth/microsoft/callback').replace(/^\/+/, '');
  if (baseUrl) {
    const callback = `${baseUrl.replace(/\/+$/, '')}${basePath ? `/${basePath}` : ''}/${redirectPath}`;
    console.log(`     ↳ register this Azure redirect URI: ${callback}`);
  }

  // 2. zrok2 installed
  console.log('\n[2] zrok2 CLI');
  const zrok = resolveZrokBin();
  if (zrok) ok(`found (${zrok})`);
  else warn('zrok2 not found — install it and `zrok2 enable <token>` (set $ZROK2_BIN if elsewhere)');

  // 3. Reserved name
  console.log('\n[3] zrok reserved name');
  const name = baseUrl ? reservedNameFromBaseUrl(baseUrl) : null;
  if (zrok && name) {
    try {
      const out = execFileSync(zrok, ['overview'], { encoding: 'utf8' });
      const line = out.split('\n').find((l) => l.includes(`${name}.`));
      if (line && /true/.test(line)) ok(`'${name}' reserved`);
      else fail(`'${name}' not reserved — run: ${zrok} create name ${name}`);
    } catch {
      warn('could not read `zrok2 overview` (is the environment enabled? `zrok2 enable <token>`)');
    }
  } else {
    warn('skipped (need zrok2 + BACKEND_BASE_URL)');
  }

  // 4. Local app
  console.log(`\n[4] Local demo app :${DEFAULT_PORT}`);
  if (await portInUse(DEFAULT_PORT)) {
    ok(`something is listening on :${DEFAULT_PORT} (start it with: npm run start:dev)`);
  } else {
    warn(`:${DEFAULT_PORT} is free — the demo app is NOT running yet (start it before/while sharing)`);
  }

  // 5. Tunnel reachable (end to end through the public URL)
  console.log('\n[5] Tunnel reachable');
  if (baseUrl) {
    const probe = webhookUrl(baseUrl);
    const shareCmd = `npm run zrok:share  (= ${zrok || 'zrok2'} share public ${DEFAULT_PORT} -n public:${name || '<name>'} --headless)`;
    try {
      const res = await fetch(probe, { method: 'POST' });
      const body = await res.text().catch(() => '');
      if (body.includes('ping')) {
        ok('tunnel + app reachable (webhook echoed validationToken) — ready');
      } else if (res.status === 502 || res.status === 503) {
        warn(`share running but the app is not listening yet (HTTP ${res.status}) — start: npm run start:dev`);
      } else if (res.status === 404) {
        warn(`reserved name resolves but NO share is running (HTTP 404) — start it: ${shareCmd}`);
      } else {
        warn(`unexpected response (HTTP ${res.status}) from ${probe} — start the share/app if needed: ${shareCmd}`);
      }
    } catch (e) {
      fail(`could not reach ${baseUrl} — is the share running? ${shareCmd}`);
    }
  } else {
    warn('skipped (no BACKEND_BASE_URL)');
  }

  console.log('\n=====================');
  if (hardFail) {
    console.log('❌ Preflight FAILED — fix the ❌ items above, then re-run.\n');
    process.exit(1);
  }
  console.log('✅ Preflight passed — start the demo and run the OAuth login flow.\n');
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
