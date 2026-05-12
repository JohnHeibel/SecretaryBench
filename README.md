# AISA Spring 26 — Internal System

An LLM benchmark that tests whether AI agents can handle **time complexity** through realistic email-thread scenarios. The agent reads a scenario's emails, then uses the tool APIs (Todos, Calendars, Emails) to complete the tasks described in those threads.

The benchmark now ships with an **MCP server** — agents call the API as MCP tools directly from Claude Code or any compatible harness. See `MCP.md` for setup.

---

## What Changed (Last Session)

- **`scenario_id` is now required on every write.** Every `POST /todos/`, `POST /calendars/{id}/events`, and `POST /emails/` must include a `scenario_id` that already exists in the store. The server returns `404` otherwise. This is the single biggest behavioral change.
- **Email API is read/write only.** There is no `DELETE /emails/{id}`. The old README listed one — it does not exist.
- **Todo updates are `PATCH`, not `PUT`.** `PATCH /todos/{id}` is a partial update; omitted fields stay unchanged. `scenario_id` and `calendar_event_id` are immutable after creation.
- **Calendar event updates are full replace.** `PUT /calendars/{id}/events/{event_id}` requires all fields, including `scenario_id`.
- **Todos can link to calendar events.** `POST /todos/` accepts an optional `calendar_event_id`; the server validates the event exists. Create the event first, then the todo.
- **Package manager is `pip` with `venv`.** See Setup section below.
- **MCP server added** (`mcp_server/`). Exposes all 22 API routes as tools. `.mcp.json` in the repo root wires it up automatically in Claude Code.

For the full, agent-facing API contract see **`docs/api_reference.md`**.

---

## Project Structure

```
.
├── app/
│   ├── main.py           # FastAPI app, error handlers, health check
│   ├── store.py          # In-memory store (resets on restart)
│   ├── models/
│   │   ├── todo.py       # TodoCreate, TodoUpdate, TodoResponse
│   │   ├── calendar.py   # CalendarCreate, CalendarResponse, EventCreate, EventResponse
│   │   └── email.py      # Email, Scenario
│   └── routers/
│       ├── todos.py
│       ├── calendar.py
│       ├── emails.py
│       └── scenarios.py
├── mcp_server/           # MCP wrapper — thin HTTP clients over FastAPI
├── engine.py             # 100-day simulation loop
├── flow_controller.py    # Scenario scheduling & chain-order enforcement
├── grader.py             # Prefix-level scoring of completed scenarios
├── loader.py             # Excel → Scenario/Email objects
├── pipeline.py           # Bridges loader objects ↔ FastAPI store
├── docs/
│   ├── api_reference.md  # Authoritative agent-facing reference
│   ├── ENGINE.md         # Simulation pipeline docs
│   ├── GRADER.md         # Scoring system docs
│   ├── FLOW_CONTROLLER.md
│   └── LOADER.md
├── tests/
│   ├── test_todos.py     # 18 tests
│   ├── test_calendars.py # 22 tests
│   ├── test_emails.py    # 8 tests
│   ├── test_scenarios.py # 18 tests
│   └── test_e2e.py       # End-to-end simulation tests
├── Emails.xlsx           # 109 scenarios across 6 sheets
├── .mcp.json             # Auto-registers MCP server in Claude Code
├── MCP.md                # MCP setup and usage guide
└── requirements.txt
```

---

## Setup

```bash
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

---

## Running

**1. Start FastAPI** (required — keep this running the whole time):

```bash
python -m uvicorn app.main:app --reload
```

Server starts at `http://127.0.0.1:8000`. Swagger UI at `/docs`, ReDoc at `/redoc`.

**2. Run the simulation:**

```bash
python engine.py                      # default: Emails.xlsx, 100 days, seed=42
python engine.py path/to/file.xlsx    # custom scenario file
```

The engine loads scenarios from the Excel file, schedules them across 100 days, serves emails to the model, and grades each scenario when its chain completes. Results print to stdout.

The simulation requires the FastAPI server to be running — the model interacts with it via MCP tools, and the grader reads back the created todos/events to score each scenario.

**3. The MCP server** starts automatically via `.mcp.json` when you open Claude Code in this repo. Verify with:

```bash
claude mcp list
# aisa: bash -c python -m mcp_server  ✓ Connected
```

See `MCP.md` for non-Claude-Code clients (MCP Inspector, Claude Desktop, custom harnesses).

---

## Running Tests

```bash
python -m pytest tests/ -v
```

| File | Tests | What it covers |
|------|-------|----------------|
| `test_todos.py` | 18 | Todo CRUD, validation, linking |
| `test_calendars.py` | 22 | Calendar & event CRUD, window validation |
| `test_emails.py` | 8 | Email read/write, ID assignment |
| `test_scenarios.py` | 18 | Scenario lifecycle, email attachment |
| `test_e2e.py` | 2 | Full simulation with stubbed AI — perfect stub scores max, bad stub scores zero |
| `test_flow_controller_pools.py` | 7 | Completed-pool retention, cross-pool accessors, delivery log + JSONL persistence |
| `test_engine_email_grading.py` | 11 | State-diff helpers, per-email "No action" override on delete, isolation from prior-email actions |

---

## Core Concepts

### `scenario_id` threads through every write

Every object the agent creates must be tagged with a `scenario_id`. The server validates the scenario exists and returns `404` if it doesn't. Load a scenario first via `POST /scenarios/` before calling any create endpoint.

```
POST /scenarios/  →  POST /emails/
                      POST /calendars/{id}/events
                      POST /todos/
```

### Linking todos to calendar events

When a task requires both a calendar event and a todo:

1. `POST /calendars/{calendar_id}/events` — get back `event_id`
2. `POST /todos/` with `calendar_event_id` set to that `event_id`

Order matters — the server validates `calendar_event_id` exists at todo creation time.

### In-memory store

All data resets when uvicorn restarts. Every benchmark run starts clean. This is intentional.

---

## API Surface

Full reference with request/response shapes, error codes, and examples: **`docs/api_reference.md`**.

Quick endpoint map:

| Group | Method | Path |
|-------|--------|------|
| Health | `GET` | `/` |
| Todos | `POST` | `/todos/` |
| | `GET` | `/todos/` |
| | `GET` | `/todos/{id}` |
| | `PATCH` | `/todos/{id}` ← partial update |
| | `DELETE` | `/todos/{id}` |
| Calendars | `POST` | `/calendars/` |
| | `GET` | `/calendars/{id}` |
| | `DELETE` | `/calendars/{id}` |
| Events | `POST` | `/calendars/{id}/events` |
| | `GET` | `/calendars/{id}/events` |
| | `GET` | `/calendars/{id}/events/{event_id}` |
| | `PUT` | `/calendars/{id}/events/{event_id}` ← full replace |
| | `DELETE` | `/calendars/{id}/events/{event_id}` |
| Emails | `GET` | `/emails/` |
| | `GET` | `/emails/{id}` |
| | `POST` | `/emails/` ← no DELETE |
| Scenarios | `GET` | `/scenarios/` |
| | `POST` | `/scenarios/` |
| | `GET` | `/scenarios/{id}` |
| | `DELETE` | `/scenarios/{id}` |
| | `POST` | `/scenarios/{id}/emails` |

### Key constraints

- `PATCH /todos/{id}` — partial update only; `scenario_id` and `calendar_event_id` cannot be changed
- `PUT /calendars/{id}/events/{event_id}` — full replace; send every field
- Calendar events must fall within the calendar's 100-day window (`start_date` through `start_date + 100 days`)
- Agent-sent email IDs: `max(all existing email_ids in store, default=0) + 1` — the counter is global, not per-scenario

### Error codes

| Status | Meaning |
|--------|---------|
| `400` | Invalid field values, constraint violation (e.g. event outside window) |
| `404` | Resource not found — also fires when `scenario_id` or `calendar_event_id` reference is missing |
| `409` | Conflict — caller-assigned ID already exists |
| `422` | Missing required field or wrong type |
| `500` | Unexpected server error |

---

## Simulation Pipeline

```
Emails.xlsx
    ↓  loader.py
Scenario/Email objects
    ↓  flow_controller.py
Scheduled across 100 days (chain-order enforced, capped to fit)
    ↓  engine.py
Daily loop:
    activate scenario → pipeline.register_scenario() seeds the store
    serve email       → model_runner.run_model_turn() calls Claude via MCP
    each email       → before/after store snapshots → grade against the diff
                      → controller.mark_served() also appends to delivery_log.jsonl
    all emails served → pipeline.fetch_scenario_results() pulls todos/events
                      → grader.define_grading_system() scores the scenario
    ↓
Aggregated results: scenario score + per-email score + by-type breakdowns + daily log
```

See `docs/ENGINE.md` for the full breakdown.

---

## Tooling

| Tool | Version / Notes |
|------|-----------------|
| Package manager | `pip` with `venv` |
| Framework | FastAPI |
| Validation | Pydantic v2 |
| Test runner | `python -m pytest` |
