#!/usr/bin/env bash
#
# SecretaryBench runner.
#
#   ./run.sh                         live benchmark — Claude Haiku, seed 42, full corpus
#   ./run.sh --model claude-sonnet-4-6 --seed 7
#   ./run.sh --limit 3               only the first 3 served emails
#   ./run.sh demo                    offline trace (oracle model, no API calls)
#   ./run.sh test                    unit test suite
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
  test)  shift; exec "$PY" -m pytest sb/tests/ -q "$@" ;;
  live)  exec "$PY" -m sb.live.runner ;;            # bare ./run.sh
  *)     exec "$PY" -m sb.live.runner "$@" ;;       # ./run.sh --model ... --seed ...
esac
