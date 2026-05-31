# SecretaryBench — AI Lane Handoff

Single primer to resume work in a fresh session. Read this + the linked docs and
you have everything. Branch: `sprint5-remediation`.

---

## 1. What this project is

SecretaryBench measures whether **a harness + a model** can run a secretary
workload (todos, calendar events, email triage) across a **100-day simulation**.
Emails are served day-by-day; the model acts via **MCP tools**; a grader scores
what it created against per-email `success_criteria`.

**Two processes:** a FastAPI store (`uvicorn app.main:app`, port 8000) and the
simulation (`python engine.py`), which drives the `claude` CLI (`claude -p`) as
the agent. The agent reaches the store through an MCP server (`mcp_server/`).

Pipeline: `engine.py → harness/ adapter → claude -p subprocess → mcp_server → FastAPI store`,
then the grader diffs the store and scores.

---

## 2. Current status (what's DONE)

The Sprint-5 remediation plan (`SPRINT5_REMEDIATION.md`) is **fully implemented —
all 13 fixes**. Full test suite: **172 passing, 0 failing** (with uvicorn up).

Highlights, all verified:
- **One runner.** The old split between `model_runner.py` and a thin adapter is
  gone. Everything lives in the **`harness/` package**: `base.py` (single source
  of truth — system prompt, MCP config, hidden tools, calendar bootstrap,
  stream-json parser with token/tool/compaction logging), `cli_base.py`
  (interface + subprocess turn loop), `claude_p.py` (claude flags), `codex.py`
  (stub). `model_runner.py` is now a thin shim. `harness.py` (flat module) deleted.
- **Calendar works on the live path** — model creates events; verified live.
- **Criteria dates resolved before grading** (FIX-2) — but see the open decision
  in §4, this has a known false-negative class.
- **token_usage.jsonl / tool_calls.jsonl** populated with cache splits; dedup intact.
- **Compaction** detected (via context-drop heuristic — claude compacts silently)
  and reported as a grading dimension (within-window vs through-compaction).
- **Failure handling**, **CLI/continuity knobs** (`--model`, `--continuity`,
  `--days`, `--limit`, env vars), **scenario_id diagnostics** at the MCP server,
  docs truth pass, response-model tightening.

**Verified live:** a 2-scenario run + a 30-email stress chain (compaction fired,
196K→18K tokens, still scored) + a **10-day / 12-scenario slice scoring 8/10**
with 0 crashes. **NOT yet run:** the full 100-day / 109-scenario live benchmark.

---

## 3. Key concepts you MUST understand (this is where confusion lives)

- **Served/simulated date = the day an email is delivered.** It is *arbitrary*
  relative to the email's content (set by the scheduler).
- **The `{date}` token in a criterion resolves to the SERVED date** — NOT the
  date the action should happen on. For "do this today" emails that's right; for
  "schedule the gala for the 3rd Friday of November" emails it's the WRONG date.
- **The resolver** (`engine._resolve_one_token`) converts tokens (`{date+2}`,
  `{date-nextweek}`, `{date-14th}`, `{date-3PM}`) into real dates, used both in
  email bodies (shown to the model) and in criteria (for grading). Tokens it
  doesn't handle (`{nextweek-wednesday}`, weekday names, ranges) stay unresolved
  → that criterion falls back to existence-only.
- **What the grader checks, by criterion prefix** (`grader.py`):
  - `TC-...` (todo): a todo **exists** (+ count, + literal content like `TC-item3`).
    **Does NOT check the due date** (known gap — see REMAINING_WORK G2).
  - `CC-...` (calendar): an event **exists** AND (if the token is a concrete date)
    the event's date matches.
  - `RS-...` (reschedule): an event exists.
  - `No action`: nothing was created.
  - anything else = free-text → **ungraded** (excluded from score, reported separately).
- **`send_email` replies are never graded** — there is no criterion prefix for them.

---

## 4. DECIDED APPROACH + NEXT ACTIONS (START HERE)

**Problem:** ~8–10 of 30 calendar scenarios currently mark a *correct* model
WRONG (15-judge audit: C19, S32, C13, C21, C11, C20, C22, C23, likely C12/S10),
because the email targets a specific date but the criterion is bare `CC-{date}`,
which resolves to the arbitrary *served/arrival* day.

**Decision made with Miguel (this is settled — do NOT re-litigate):**
- Fix the **answer key** (`Emails.xlsx`) **by hand** where a criterion has the
  wrong/missing date. Excel is editable by Miguel, NOT by code or by you.
- The **code** must faithfully grade whatever date the criterion specifies.
- **NO code workarounds** that paper over bad criteria (no "make bare `{date}`
  lenient", no "grade against the email body"). Bare `{date}` stays **strict =
  served day**; if a criterion shouldn't be the arrival day, Miguel relabels it.

**For the code to "pick up" a proper date, two gaps must close (both code):**
1. The **resolver** must read the date forms the criteria use. It already reads
   `{date}`, `{date+N}`, `{date-nextweek}`, `{date-Nth}`, `{date-3PM}`,
   `{nextweek-date+N}`. It does NOT yet read: **weekday tokens**
   (`{date-Tuesday}`, `{nextweek-wednesday}`, `{nextweek-friday}`),
   **Nth-weekday-of-month** (`{third Friday}` …), and **date+time combos**
   (`{date+1- 3PM}`, `{date- this week 11pm}`).
2. **TC (todo) criteria don't check the due date at all** — must mirror CC.

**Criterion-token groups (full scan of Emails.xlsx success-criteria column):**
- **Group 1 — already date-graded:** `{date}` (46×, = served day) + ~20 specific
  tokens. Working.
- **Group 2 — proper date IS in the criterion but the resolver can't read it**
  (so it grades existence-only today): `{nextweek-wednesday}`×2, `{nextweek-friday}`,
  `{third Friday}`, `{third Friday -1 dinner}`, `{date+1- 3PM}`×2,
  `{date+2- 11AM}`×2, `{date- this week 11pm}`, `{nextweek-Thursday, 3pm GMT}`.
  **Fixing the resolver grades these correctly with ZERO Excel edits** — this is
  the bulk of the win and it's pure code.
- **Group 3 — not dates:** `{greenlight product A}` (S21 — it's a *content* check;
  should be unbraced `TC-greenlight`), `{deadline}` (T08 — "flag ambiguity", a
  judgment the structural grader can't verify), `{C}`×4 (S01 — emails say abstract
  "date A/B/C" with no real date → underspecified; existence-only unless emails
  are rewritten). Rule of thumb for Miguel: **`{...}` = check a date; plain text
  after the prefix = check the todo/event text contains it.**

**NEXT ACTIONS, in order:**
1. **CODE — extend the resolver** (`engine._resolve_one_token`): weekday tokens
   (`{date-<weekday>}` = next future occurrence on/after served date;
   `{nextweek-<weekday>}` = that weekday of next week), Nth-weekday-of-month
   (`{third Friday}` / `{3rd Friday of November}`), and the date+time combos.
   Semantics = "next future occurrence", consistent with existing `{date-Nth}`.
   + tests. → recovers Group 2.
2. **CODE — add TC due-date matching** in `grader.py` (mirror `_event_matches_date`):
   a TC criterion whose token resolves to a concrete date requires a todo whose
   `due_date` matches; bare/unresolvable → existence-only. + tests.
3. **CODE — write `docs/TOKEN_REFERENCE.md`**: every token the code understands,
   so Miguel's Excel edits use valid forms.
4. **RUN** the 10-day slice (`python engine.py Emails.xlsx --days 10 --limit 12`,
   server up) and/or the full run; identify which scenarios STILL score wrong.
5. **HAND MIGUEL a short, exact Excel checklist** — only the still-broken few,
   as "scenario X: change `CC-{date}` → `CC-{…}`". Known trivial one to include:
   **S21: `TC-{greenlight product A}` → `TC-greenlight`.**

The remaining bare-`{date}` cases where the email means a different date (C19
"3rd Friday of Nov", S32 "first Wed Aug", C20 "two weeks from Monday", C22, C23,
C11) are **Miguel's manual edits AFTER step 1** gives him tokens — several may be
hard to express; decide per-scenario after the re-run, don't pre-solve them.

**Important:** Miguel was getting overwhelmed by the token zoo. Keep his surface
area tiny — do all the code, re-run, then give him a SHORT certain list. The weird
tokens are pre-existing dataset content, not bugs you introduced.

---

## 5. What else is left (full list in `docs/REMAINING_WORK.md`)

Covered by the §4 next-actions:
- **G5** resolver gaps (weekday / Nth-weekday / date+time) — recovers Group 2, no Excel edit.
- **G2** TC due-date never checked — 21 of 48 todo criteria specify a deadline the grader ignores.
- **G1** is resolved by the §4 decision (fix ground truth + faithful code; NOT a leniency hack).

Lower priority:
- **H3** `context_window_exceeded` always False; **V2** test hygiene (test_pipeline runs sims at import; no pytest-timeout).

Dataset-bound (Miguel edits Excel; code can't invent the data):
- **G3** email replies unscored (no criterion type exists for them); **G4** 3 "delete/create meeting" free-text criteria ungraded; Group-3 `{deadline}`/`{C}` are underspecified scenarios.

Out of scope / unverified:
- **H1** second harness (Codex stub); **H4** OpenRouter path wired but never run;
- **V1** the full 100-day live run has not been executed.

Do NOT "fix": stream-json dedup, `resume_session` no-op, day-100 clamping — all correct.

---

## 6. How to run

```bash
source .venv/bin/activate                       # NOT venv/ — repo uses .venv
python -m uvicorn app.main:app --port 8000      # terminal A (leave running)

# Offline tests (no server, no claude):
python -m pytest tests/ -q -k "not test_perfect_stub and not test_bad_stub" -o addopts=""
# Full suite (needs uvicorn up): 172 should pass
python -m pytest tests/ -q -o addopts=""

# Live runs (need uvicorn + authenticated `claude`):
python engine.py Emails.xlsx --days 10 --limit 12   # bounded smoke (what we ran)
python engine.py Emails.xlsx                         # FULL 100-day run (not yet done)

# Stress chain (prove compaction):
python tests/generate_stress_chain.py -o /tmp/stress.xlsx --emails 30
python engine.py /tmp/stress.xlsx
```

Useful flags: `--model`, `--reasoning {low,medium,high}`, `--continuity {0,1}`,
`--days N`, `--limit N` (first N scenarios — `--days` alone does NOT reduce turns,
the scheduler spreads all scenarios across the days; use `--limit` to bound).
Logs: `token_usage.jsonl`, `tool_calls.jsonl`, `delivery_log.jsonl`.

---

## 7. Hard constraints / conventions

- **Excel edits are Miguel's job, done by hand** — only he edits `Emails.xlsx`
  success criteria. Code must never auto-edit the sheet or hack around bad
  criteria. The division of labor: Miguel fixes the answer key; code grades it
  faithfully. (This replaces the earlier "Excel is frozen" framing — Miguel will
  relabel the few wrong criteria himself, AFTER the code can read proper tokens.)
- **Goal:** the score must MEAN something — fixing a thing is only "correct" if it
  makes the score reflect real secretary competence. Don't optimize away the token
  cost of long chains (that cost is the benchmark); do remove false grades.
- venv is `.venv/`. Run from repo root. macOS (no GNU `timeout`; no `pip` in venv).

---

## 8. File map

- `harness/` — the runner core (base/cli_base/claude_p/codex). **Single source of truth.**
- `engine.py` — simulation loop, token resolution, grading call sites, CLI.
- `grader.py` — `define_grading_system`, `_check_criteria`, splitter, date/content matching.
- `flow_controller.py` — scheduling, delivery_log (now records failures).
- `mcp_server/server.py` — MCP tool surface + scenario_id diagnostics.
- `pipeline.py` — loader objects ↔ FastAPI store over HTTP.
- `docs/` — `GRADER.md`, `REMAINING_WORK.md` (the detailed backlog), `MCP.md`, `api_reference.md`.
- `SPRINT5_REMEDIATION.md` — the original (now-completed) plan.
- Tests: `tests/test_fix2..fix12_*.py` (new, per-fix), `test_e2e.py`, `test_pipeline.py`.
