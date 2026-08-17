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
# NOTE: build the venv with an explicit 3.13, not bare `python3`. macOS ships 3.9,
# which cannot install `mcp` and cannot evaluate the PEP 604 annotations in
# sb/live/mcp_app.py — the failure surfaces much later as "mcp server did not come up".
SETUP="  uv venv --python 3.13 .venv && uv pip install -r requirements.txt --python .venv/bin/python"
if [[ ! -x "$PY" ]]; then
  echo "error: $PY not found. Create the venv first (Python 3.10+ required):" >&2
  echo "$SETUP" >&2
  exit 1
fi
if ! "$PY" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)'; then
  echo "error: $PY is $("$PY" -V 2>&1), but the harness needs Python 3.10+." >&2
  echo "Delete .venv and rebuild it (a venv cannot be upgraded in place):" >&2
  echo "$SETUP" >&2
  exit 1
fi

cmd="${1:-live}"
case "$cmd" in
  demo)  shift; exec "$PY" -m sb.demo "$@" ;;
  test)  shift; exec "$PY" -m pytest sb/tests/ -q "$@" ;;
  # `shift || true` because a bare `./run.sh` has no args to shift (set -e would abort).
  live)  shift || true; exec "$PY" -m sb.live.runner "$@" ;;   # ./run.sh [live] --model ...
  *)     exec "$PY" -m sb.live.runner "$@" ;;       # ./run.sh --model ... --seed ...
esac
