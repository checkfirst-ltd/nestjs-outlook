#!/usr/bin/env bash
#
# Start a zrok tunnel that bridges the fixed reserved name -> this demo's local port.
# Microsoft Graph needs a public HTTPS URL to (a) redirect the OAuth callback back and
# (b) deliver webhook notifications; localhost:8888 is not reachable from Microsoft.
#
# The public URL stays FIXED as https://<name>.shares.zrok.io across every run, so you
# register it in Azure once. See ZROK.md for the full setup.
#
# Usage:
#   npm run zrok:share                # bridges :8888 -> nestjs-outlook-demo
#   ZROK_PORT=9000 npm run zrok:share # override the local port
#   ZROK_NAME=my-name npm run zrok:share
#
# One-time prerequisites (see ZROK.md):
#   zrok2 enable <account-token>
#   zrok2 create name nestjs-outlook-demo
set -euo pipefail

ZROK_PORT="${ZROK_PORT:-${PORT:-8888}}"
ZROK_NAME="${ZROK_NAME:-nestjs-outlook-demo}"
ZROK_NAMESPACE="${ZROK_NAMESPACE:-public}"

# Resolve the zrok v2 binary the same way the preflight does.
resolve_zrok() {
  for bin in "${ZROK2_BIN:-}" zrok2 "$HOME/bin/zrok2"; do
    [ -z "$bin" ] && continue
    if "$bin" version >/dev/null 2>&1; then
      echo "$bin"
      return 0
    fi
  done
  return 1
}

if ! ZROK_BIN="$(resolve_zrok)"; then
  echo "❌ zrok2 not found. Install zrok v2, then: zrok2 enable <account-token>" >&2
  echo "   (set \$ZROK2_BIN if the binary lives somewhere non-standard)" >&2
  exit 1
fi

echo "▶ Bridging https://${ZROK_NAME}.shares.zrok.io  ->  http://localhost:${ZROK_PORT}"
echo "  (start the demo with: npm run start:dev — it must listen on :${ZROK_PORT})"
echo

exec "$ZROK_BIN" share public "$ZROK_PORT" -n "${ZROK_NAMESPACE}:${ZROK_NAME}" --headless
