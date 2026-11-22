#!/usr/bin/env bash
# Wrapper used by cron: loads ~/.ha_trader/.env, runs ato_trader.py, appends to a log.
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
HOME_DIR="${HA_TRADER_HOME:-$HOME/.ha_trader}"
set -a; . "$HOME_DIR/.env"; set +a
mkdir -p "$HOME_DIR/logs"
exec "$HERE/../.venv/bin/python" "$HERE/ato_trader.py" "$@" \
  >> "$HOME_DIR/logs/$(date +%Y-%m).log" 2>&1
