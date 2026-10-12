#!/usr/bin/env bash
# Temporary public URL via Cloudflare quick tunnel (dies when processes stop).
set -euo pipefail
ROOT="$(cd "$(dirname "$0")" && pwd)"
DEPLOY="/workspace/hye-rate-quote-deploy"
PORT=8765

if [[ "${1:-}" == "--parse" ]]; then
  python3 "$ROOT/scripts/parse_rates.py"
fi

rm -rf "$DEPLOY"
mkdir -p "$DEPLOY/data/parsed" "$DEPLOY/static"
cp "$ROOT/index.html" "$DEPLOY/"
cp -r "$ROOT/static/." "$DEPLOY/static/"
cp "$ROOT/data/parsed/"*.json "$DEPLOY/data/parsed/" 2>/dev/null || true
# rates.json unused by UI and large
rm -f "$DEPLOY/data/parsed/rates.json"

pkill -f "python3 -m http.server ${PORT}" 2>/dev/null || true
pkill -f "cloudflared tunnel --url http://127.0.0.1:${PORT}" 2>/dev/null || true
sleep 0.5

CFBIN="${CLOUDFLARED_BIN:-/tmp/cloudflared}"
if [[ ! -x "$CFBIN" ]]; then
  curl -sL -o "$CFBIN" "https://github.com/cloudflare/cloudflared/releases/download/2026.10.0/cloudflared-linux-amd64"
  chmod +x "$CFBIN"
fi

nohup python3 -m http.server "$PORT" --bind 127.0.0.1 --directory "$DEPLOY" >/tmp/hye-http.log 2>&1 &
echo $! >/tmp/hye-http.pid
: >/tmp/cloudflared.log
nohup "$CFBIN" tunnel --url "http://127.0.0.1:${PORT}" --no-autoupdate >/tmp/cloudflared.log 2>&1 &
echo $! >/tmp/cloudflared.pid

for i in $(seq 1 40); do
  URL=$(rg -o 'https://[a-zA-Z0-9-]+\.trycloudflare\.com' /tmp/cloudflared.log 2>/dev/null | tail -1 || true)
  if [[ -n "${URL:-}" ]]; then
    echo "PUBLIC_URL=$URL"
    echo "TEMP: URL dies when cloudflared/http.server stop or the box restarts."
    exit 0
  fi
  sleep 1
done
echo "Failed to obtain trycloudflare URL; see /tmp/cloudflared.log" >&2
exit 1
