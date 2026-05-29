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

### Examples

```bash
python engine.py --model claude-sonnet-4-6 --reasoning high
python engine.py mydata.xlsx --harness claude-code
OPENROUTER_API_KEY=sk-... python engine.py --openrouter --api-base https://openrouter.ai/api/v1
```

### Environment-variable knobs

Read by `model_runner.py` / the adapter:

| Variable | Default | Meaning |
|---|---|---|
| `FASTAPI_BASE_URL` | `http://localhost:8000` | Where the agent reaches the API |
| `CONVERSATION_CONTINUITY` | `1` | `0` = fresh `claude` session per email instead of resuming |
| `TOKEN_LOG_PATH` | `token_usage.jsonl` | Per-round token usage log |
| `TOOL_LOG_PATH` | `tool_calls.jsonl` | Per-call tool log |

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

- **Adapter path vs. model_runner.** `engine.py` prefers the `harness.py` adapter
  (it tries `get_adapter()` first) over `model_runner.py`. Both build nearly
  identical `claude` commands, but the richer token/tool logging
  (`token_usage.jsonl`, `tool_calls.jsonl`) lives in `model_runner.py`, which the
  default adapter path does **not** trigger.
- **MCP config is injected inline.** The adapter passes its own `--mcp-config`, so
  you do **not** need `.mcp.json` registered in Claude Code for the simulation to
  work.
- **Tests:** `python -m pytest tests/ -v`.

See `README.md` for the full project overview and `MCP.md` for using the MCP
server with non-Claude-Code clients.
