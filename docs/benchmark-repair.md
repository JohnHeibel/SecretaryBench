# SecretaryBench repair register

The single living record of what is wrong with this benchmark, what we changed, and
what has actually been proven. Read this before starting work; update it in the same
turn as any code change.

**Status legend:** `open` → `investigating` → `fix proposed` → `applied` → `verified`.
Nothing reaches `verified` without a named artifact (a passing test, a re-scored run,
a diff). IDs are permanent and never reused.

---

## Dashboard

| ID | problem | severity | status | phase |
|---|---|---|---|---|
| E-1 | venv built on macOS system Python 3.9; `mcp` and `mcp_app.py` need 3.10+ | blocks-running | **verified** | 0 |
| E-2 | `mcp>=1.0.0` unbounded resolves to 2.0.0, which moved `FastMCP`; tool server never starts | blocks-running | **verified** | 0 |
| E-3 | `start_mcp` discarded stderr, so any startup crash read as "did not come up" | slows-work | **verified** | 0 |
| M-1 | `run.sh` `live` branch dropped every CLI argument, silently running defaults | blocks-measurement | **verified** | 0 |
| M-2 | runner never reported which model the CLI actually served | blocks-measurement | **verified** | 0 |
| M-3 | `--model` behaviour across `--resume` turns unverified on CLI 2.1.233 | distorts-measurement | **verified** | 0 |
| M-4 | documented model roster may be stale | cosmetic | **closed** | 0 |
| M-5 | runner inherits `CLAUDE_CODE_*` env when launched from a Claude Code session | unknown | open | 0 |
| G-* | grader identity: object matching, kind filter, description in haystack | blocks-measurement | open | 2 |
| A-* | attribution and no-action grading hang on model-supplied `email_id` | distorts-measurement | open | 3 |
| O-* | no machine-readable output, no store dump, lossy tool trace | blocks-iteration | open | 1 |
| C-* | config and provenance not stamped; durable record stale | distorts-measurement | partly done | 5 |
| K-* | corpus authoring drift (free-typed dates, malformed tokens, prose/answer mismatch) | distorts-measurement | open | 4 |
| V-* | retrieval span untested; tier data unread by `sb.analyze` | blocks-claims | open | 5 |

`G/A/O/C/K/V` are placeholders. They get filled in by the category fan-out (phase A)
before any of phases 1–5 begin.

---

## Phase 0 — model resolution and environment (complete)

### E-1 · venv on the wrong Python
**Status:** verified.
`/usr/bin/python3` on this machine is 3.9.6. `mcp` requires 3.10+, and
`sb/live/mcp_app.py:50,81` use PEP 604 (`str | None`) in FastMCP tool signatures, which
FastMCP evaluates at runtime — a hard TypeError on 3.9. The documented setup line
(`run.sh:18`, `BENCHMARK_RESULTS.md` §4) says bare `python3`, which builds a 3.9 venv.
**Fix:** build with an explicit interpreter — `uv venv --python 3.13 .venv`.
**Verified by:** `.venv/bin/python -m pytest sb/tests -q` → 62 passed.
**Still to do:** correct the setup line in `run.sh` and the docs so the next person on a
Mac does not hit this; consider a version guard beside the existing venv check.

### E-2 · `mcp` 2.0 broke the tool server
**Status:** verified.
`requirements.txt` pinned `mcp>=1.0.0` with no upper bound. The MCP Python SDK shipped
2.0.0, which renamed `FastMCP` to `MCPServer` and moved it from `mcp.server.fastmcp` to
`mcp.server.mcpserver`. A fresh install therefore fails at
`sb/live/mcp_app.py:25` with `ModuleNotFoundError: No module named 'mcp.server.fastmcp'`,
the MCP server never binds, and the runner aborts before reaching any model.
**Fix:** pin `mcp>=1.0.0,<2.0.0` (resolves to 1.29.0).
**Verified by:** `from mcp.server.fastmcp import FastMCP` imports; live smoke reaches the
model and the model makes real MCP tool calls.
**Deliberately not done:** porting `mcp_app.py` to the 2.0 `MCPServer` API. That would
change the tool surface the model under test sees, which is a benchmark change and needs
its own verification, not a dependency bump.

### E-3 · MCP startup failures were unreadable
**Status:** verified.
`start_mcp` passed `stderr=subprocess.DEVNULL`, so E-2's traceback surfaced only as
`RuntimeError: mcp server did not come up`.
**Fix:** keep stderr, poll `proc.poll()`, and put the child's traceback in the raised
error.
**Verified by:** this is how E-2 was diagnosed.

### M-1 · `run.sh` silently discarded every flag
**Status:** verified. **This is the reported bug.**
`run.sh:26` read `live) exec "$PY" -m sb.live.runner ;;` — no `shift`, no `"$@"`. The
`demo` and `test` branches both forward arguments; `live` did not. So
`./run.sh live --model claude-opus-5 --seed 7` executed the runner with **zero
arguments** and fell back to argparse defaults at `runner.py:499-514`:
`--model claude-haiku-4-5`, `--seed 42`, `--days 60`, `--daily-max 5`,
`--corpus corpus`. That is exactly the report: an unknown default model regardless of
the flag, and an identical configuration — hence an identical score — every run.
The trap only fires on `./run.sh live --model X`; `./run.sh --model X` falls through to
the `*)` branch and always worked, which is why it went unnoticed.
**Fix:** `live) shift || true; exec "$PY" -m sb.live.runner "$@" ;;`. The `|| true` is
load-bearing: a bare `./run.sh` has nothing to shift and `set -e` would abort.
**Verified by:** `./run.sh live --help` now reaches argparse; two live smokes launched
via `./run.sh live --model X` served two different models (see M-3).

### M-2 · the runner never reported the model actually served
**Status:** verified.
The header printed the model that was *requested*. The `system`/`init` event in the
CLI's stream already carries the model that was *loaded*, and `_parse_stream` read that
event but kept only `session_id`. A silent substitution was therefore unobservable, in
the live output and in every saved log in `outputs/` and `past/`.
**Fix:** `_parse_stream` and `_parse_codex` now return a 5th value, the resolved model.
The runner prints it once, flags a mismatch in red, warns on mid-run drift (M-3), and
stamps the served model plus seed, levers, and corpus next to the final score so a saved
log cannot be mislabelled.
**Verified by:** parser unit-checked on synthetic streams for both drivers; live smokes
print `model served claude-opus-4-8 ✓` and `claude-sonnet-4-5-20250929 ✓`.
Note the match test is `startswith`, because the CLI resolves an alias to a dated
snapshot (`claude-sonnet-4-5` → `claude-sonnet-4-5-20250929`).

### M-3 · `--model` across `--resume`
**Status:** verified.
A one-turn smoke cannot see this: if `--model` were ignored on resume, day 1 would be
correct and every later day would drift to the account default, producing exactly the
"same result every time" symptom even with M-1 fixed.
**Fix:** the resolved model is checked on **every** turn, not just the first, and any
change prints a `MODEL DRIFT` warning naming the day.
**Verified by:** a 3-day / 9-email smoke (`--limit 9`, two `--resume` turns) served
`claude-opus-4-8` on every turn with no drift warning. `--model` does survive `--resume`
on CLI 2.1.233, so M-1 was the whole story.

### M-4 · roster staleness
**Status:** closed, not a problem. `claude-opus-4-8`, `claude-sonnet-4-5` and
`claude-haiku-4-5` are all current, active model IDs. The roster in
`BENCHMARK_RESULTS.md` §2 resolves fine; the OpenAI half is untestable here only because
`codex` is not installed on this machine.

### M-5 · inherited environment
**Status:** open, unquantified.
`runner.py` spawns the CLI with the parent environment. A shell inside a Claude Code
session carries `CLAUDECODE`, `CLAUDE_CODE_SESSION_ID`, `CLAUDE_CODE_ENTRYPOINT` and
`CLAUDE_CODE_CHILD_SESSION`. Also `~/.claude/settings.json` sets `"model": "opus[1m]"`
globally, so any path where `--model` fails to reach the CLI inherits Opus rather than
erroring. Not observed to cause a problem — the smokes served the requested models — but
runs launched from inside an agent session are not obviously equivalent to runs launched
from a plain shell. Worth settling before the phase 6 comparison run.

---

## Corpus health check (phase 0 side-effect)

`.venv/bin/python -m sb.scale --filler 0 --seed 42 --days 200 --dst build/scaled0`

```
scaled corpus: 167 emails, 167 served over 57 days
filler: 0 junk emails burying 24 authored needle(s)
needle span: max 83, mean 31.6  (n=24)
oracle: 167/167 = 100%
```

The corpus lints clean and every answer key is satisfiable. The 57-day span matches the
`outputs/opus.md` and `outputs/sonnet.md` logs exactly, confirming those ran at default
levers (`daily_max=5`) rather than the `daily_max=21` pinned in `BENCHMARK_RESULTS.md`.
Mean needle span of 31.6 is low, which is the V-category concern: the retrieval axis the
benchmark is built to measure is barely exercised by the authored corpus alone.

---

## Open question inherited from before phase 0

The reported score was "about 51%". No committed log shows 51% — `outputs/opus.md` is
54%, `outputs/sonnet.md` 54%, `past/claude-haiku-4-5.md` 59%. So at least one run exists
whose artifact was never saved. If that log survives anywhere it is worth recovering:
its header states the model the runner *thought* it was using, which would confirm M-1
directly rather than by inference.

---

## Changelog

- **2026-08-17** — Phase 0. Fixed E-1 (venv on 3.13), E-2 (`mcp<2.0.0`), E-3 (MCP stderr),
  M-1 (`run.sh` argument forwarding), M-2 (resolved-model reporting + config stamp),
  M-3 (per-turn drift detection). Closed M-4. Opened M-5. Confirmed corpus oracle 100%.
  62 unit tests pass. Confirmed corpus oracle 100%.

  Live smokes, all launched via the previously-broken `./run.sh live --model X` form:

  | smoke | requested | served | result |
  |---|---|---|---|
  | 3 emails, 1 turn | `claude-sonnet-4-5` | `claude-sonnet-4-5-20250929` ✓ | 3/3 |
  | 3 emails, 1 turn | `claude-opus-4-8` | `claude-opus-4-8` ✓ | 3/3 |
  | 9 emails, 3 days, 2 resumes | `claude-opus-4-8` | `claude-opus-4-8` ✓ no drift | 9/9 |

  Two *different* models were served, and neither is the silent default
  (`claude-haiku-4-5`) — chosen deliberately so a persisting M-1 could not hide.
  The 9/9 and 3/3 scores are not a capability signal: these are the first easy emails
  in the plan, and the grader problems (G-*) live further into the run.
