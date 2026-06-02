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
    PORT="${PORT:-8099}"
    url="http://localhost:$PORT"

    # don't launch behind a process that already owns the port — point at it instead.
    if lsof -nP -iTCP:"$PORT" -sTCP:LISTEN >/dev/null 2>&1; then
      echo "error: port $PORT is already in use." >&2
      echo "  if that's an old editor, stop it (pkill -f authoring.server) and retry;" >&2
      echo "  otherwise open $url — it may already be serving." >&2
      exit 1
    fi

    if [[ ! -f authoring/web/dist/index.html ]]; then
      echo "building scenario editor (first run)…" >&2
      ( cd authoring/web && npm install --silent && npm run build )
    fi

    # A detached watcher announces the URL only once the server actually answers,
    # so we never report "ready" while the browser would still see "connection refused".
    # It self-terminates, so there's nothing to clean up.
    ( printf 'starting scenario editor' >&2
      for _ in $(seq 1 30); do
        if curl -fsS -o /dev/null "$url" 2>/dev/null; then
          echo >&2; echo "scenario editor → $url  (Ctrl+C to stop)" >&2; exit 0
        fi
        printf '.' >&2; sleep 1
      done
      echo >&2; echo "warning: editor still not answering after 30s — see any errors above." >&2 ) &

    # exec so uvicorn BECOMES this process: Ctrl+C goes straight to it (clean shutdown,
    # no orphaned server), and its startup errors print live to the terminal.
    exec env PORT="$PORT" "$PY" -m authoring.server ;;
  ui-dev)
    # hot-reload dev: API on :8099, Vite dev server (proxying /api) on :5173.
    "$PY" -m authoring.server &
    api_pid=$!
    trap 'kill "$api_pid" 2>/dev/null' EXIT
    ( cd authoring/web && npm install --silent && npm run dev ) ;;
  live)  exec "$PY" -m sb.live.runner ;;            # bare ./run.sh
  *)     exec "$PY" -m sb.live.runner "$@" ;;       # ./run.sh --model ... --seed ...
esac
