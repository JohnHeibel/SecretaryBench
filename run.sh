#!/usr/bin/env bash
#
# SecretaryBench runner.
#
#   ./run.sh                         live benchmark — Claude Haiku, seed 42, full corpus
#   ./run.sh --model claude-sonnet-4-6 --seed 7
#   ./run.sh --limit 3               only the first 3 served emails
#   ./run.sh demo                    offline trace (oracle model, no API calls)
#   ./run.sh test                    unit test suite
#   ./run.sh ui                      scenario editor — http://localhost:8099 (builds web if needed)
#   ./run.sh ui-dev                  scenario editor with hot-reload (Vite :5173 + API :8099)
#
# The live run starts and stops its own state server — nothing else to launch.
set -euo pipefail
cd "$(dirname "$0")"

PY=".venv/bin/python"
if [[ ! -x "$PY" ]]; then
  echo "error: $PY not found. Create the venv first:" >&2
  echo "  python3 -m venv .venv && .venv/bin/pip install -r requirements.txt" >&2
  exit 1
fi

cmd="${1:-live}"
case "$cmd" in
  demo)  shift; exec "$PY" -m sb.demo "$@" ;;
  test)  shift; exec "$PY" -m pytest sb/tests/ authoring/ -q "$@" ;;
  ui)
    # one-line scenario editor: build the web app if missing, then serve it + the API.
    if [[ ! -f authoring/web/dist/index.html ]]; then
      echo "building scenario editor (first run)…" >&2
      ( cd authoring/web && npm install --silent && npm run build )
    fi
    echo "scenario editor → http://localhost:8099" >&2
    exec "$PY" -m authoring.server ;;
  ui-dev)
    # hot-reload dev: API on :8099, Vite dev server (proxying /api) on :5173.
    "$PY" -m authoring.server &
    api_pid=$!
    trap 'kill "$api_pid" 2>/dev/null' EXIT
    ( cd authoring/web && npm install --silent && npm run dev ) ;;
  live)  exec "$PY" -m sb.live.runner ;;            # bare ./run.sh
  *)     exec "$PY" -m sb.live.runner "$@" ;;       # ./run.sh --model ... --seed ...
esac
