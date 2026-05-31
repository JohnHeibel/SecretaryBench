# Getting Started

How to run SecretaryBench end-to-end. This is a **two-process** benchmark: one
process runs the API server, the other runs the simulation, which drives the
`claude` CLI as the agent.

---

## Prerequisites

- **Python 3.10+** (`python3` is fine — the venv step below provides the `python`
  binary the MCP server needs).
- **`claude` CLI installed and authenticated.** Check with `claude --version`.
  If you're not logged in, run `claude` once interactively to authenticate (or
  use the OpenRouter flags below).

---

## 1. One-time setup

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

> **Always run from inside the activated venv.** The agent's MCP server is
> spawned by `claude` as `bash -c "python -m mcp_server"`. Many machines have
> only `python3`, not `python` — activating the venv puts a `python` symlink on
> PATH, which the `claude` subprocess inherits. Skip this and the agent will
> have no tools and every scenario scores 0.

---

## 2. Start the API server (Terminal A — leave running)

```bash
source venv/bin/activate
python -m uvicorn app.main:app --reload
```

Server starts at `http://127.0.0.1:8000`. Swagger UI at `/docs`, ReDoc at
`/redoc`. The in-memory store resets every time uvicorn restarts — that's
intentional, so every run starts clean.

---

## 3. Run the simulation (Terminal B)

```bash
source venv/bin/activate
python engine.py                  # defaults: Emails.xlsx, 100 days, seed 42, haiku
```

The engine loads scenarios, schedules them across 100 days, serves emails to the
model via `claude`, and grades each scenario as its chain completes. Results
print to stdout.

---

## CLI flags (`engine.py`)

| Flag | Default | Meaning |
|---|---|---|
| `path` (positional) | `Emails.xlsx` | Scenario workbook to load |
| `--harness` | `claude-code` | Adapter. Only `claude-code` works; `codex` is a stub that raises `NotImplementedError` |
| `--model` | `claude-haiku-4-5` | Model passed to `claude --model` |
| `--reasoning` | none | `low` / `medium` / `high` → `claude --effort` |
| `--openrouter` | off | If set, passes `$OPENROUTER_API_KEY` as `ANTHROPIC_API_KEY` to the subprocess |
| `--api-base` | none | Sets `ANTHROPIC_BASE_URL` (for OpenRouter / custom endpoint) |
| `--continuity` | env | `1` resumes the per-scenario session; `0` runs a fresh session per email. Defaults to `CONVERSATION_CONTINUITY` |

### Examples

```bash
python engine.py --model claude-sonnet-4-6 --reasoning high
python engine.py mydata.xlsx --harness claude-code
python engine.py --continuity 0          # A/B: fresh session per email
OPENROUTER_API_KEY=sk-... python engine.py --openrouter --api-base https://openrouter.ai/api/v1
```

### Environment-variable knobs

All of these are read once by `harness/base.py` (the shared core that both the
adapter and the `model_runner` shim use), so they apply to the live `claude -p`
path. `CLAUDE_MODEL` / `CLAUDE_REASONING` are the env equivalents of `--model` /
`--reasoning` (the flag wins if both are set).

| Variable | Default | Meaning |
|---|---|---|
| `FASTAPI_BASE_URL` | `http://localhost:8000` | Where the agent reaches the API |
| `CLAUDE_MODEL` | `claude-haiku-4-5` | Default model (same as `--model`) |
| `CLAUDE_REASONING` | none | Default reasoning effort → `claude --effort` |
| `CONVERSATION_CONTINUITY` | `1` | `0` = fresh `claude` session per email instead of resuming. Overridable with `--continuity` |
| `TOKEN_LOG_PATH` | `token_usage.jsonl` | Per-round token usage log |
| `TOOL_LOG_PATH` | `tool_calls.jsonl` | Per-call tool log |
| `HARNESS_RETRY_ON_FAILURE` | `0` | `1` = retry a crashed/timed-out turn once before recording the failure |

---

## Will it work? — checklist

Before running, make sure all of these are true:

- [ ] **venv created and activated** (provides deps *and* the `python` binary the
      MCP server is spawned with).
- [ ] **`pip install -r requirements.txt` completed** — `engine.py` fails at
      import otherwise (e.g. missing `openpyxl`).
- [ ] **`uvicorn` running** in a separate terminal — the agent and grader both
      talk to it over HTTP.
- [ ] **`claude` CLI authenticated** — the subprocess runs
      `claude -p ... --permission-mode bypassPermissions` using your existing
      login. Or supply `--openrouter` / `--api-base` with a key.

---

## Good to know

- **One runner, one behavior.** The "drive `claude` for one email" logic lives in
  the `harness/` package (`harness/base.py` is the single source of truth for the
  system prompt, MCP config, calendar bootstrap, and the stream-json parser that
  writes `token_usage.jsonl` / `tool_calls.jsonl` and detects compaction).
  `engine.py` uses `get_adapter()`; `model_runner.py` is now a thin shim over the
  same core. Token/tool/compaction logging fires on the default `claude -p` run —
  there is no longer a logging-vs-no-logging split between two runners.
- **MCP config is injected inline.** The adapter passes its own `--mcp-config`, so
  you do **not** need `.mcp.json` registered in Claude Code for the simulation to
  work.
- **Tests:** `python -m pytest tests/ -v` (needs `uvicorn` running for the
  end-to-end tests).

See `README.md` for the full project overview and `docs/MCP.md` for using the MCP
server with non-Claude-Code clients.
