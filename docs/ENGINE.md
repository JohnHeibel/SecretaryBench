# engine.py — Pipeline & Time Simulation

`engine.py` is the heartbeat of the 100-day benchmark. It drives the day loop, resolves date placeholders inside email content, hands emails to the model (via `model_runner.py` or a test stub), and connects to the grader after each scenario completes.

It coordinates four other modules:
- **`loader.py`** — reads the Excel file into `Scenario`/`Email` objects
- **`flow_controller.py`** — decides which emails are due each day, enforces chain order, caps at day 100
- **`pipeline.py`** — bridges loader objects and the FastAPI store (seeding + fetching results)
- **`grader.py`** — scores completed scenarios

---

## How to run

```bash
python3 engine.py                      # default: Emails.xlsx
python3 engine.py path/to/file.xlsx    # custom path
```

Default settings: 100-day simulation starting January 1, 2000, seeded with `seed=42` for reproducibility.

---

## What it does, top-down

```
load_scenarios("Emails.xlsx")                 # 109 Scenario objects
controller = FlowController(scenarios)        # 109 in inactive_pool
controller.build_schedule(total_days=100)

for day in 0..99:
    ready = controller.step(day)              # emails due today (chain-order enforced)

    for scenario in controller.newly_activated_today():
        register_scenario(scenario, sim_date) # seed the FastAPI store

    for email, scenario, idx in ready:
        resolved = apply_date_substitutions(email, sim_date)   # tokens → dates
        before = fetch_scenario_results(scenario)              # snapshot store
        run_model_turn(resolved, sim_date)                     # Claude via MCP
        after = fetch_scenario_results(scenario)               # snapshot again
        # per-email grade against the diff (skipped if email has no criteria)
        controller.mark_served(scenario.scenario_id, idx, day=day, sim_date=...)

    for scenario in controller.completed_scenarios_today():
        state = fetch_scenario_results(scenario)               # full cumulative state
        result = define_grading_system(scenario, state["calendar"], state["todos"])
        # accumulate scenario-level score

    sim_date += 1 day

return aggregated results
```

Per-day, the engine: (1) asks the controller what's ready, (2) seeds the store for any newly activated scenarios, (3) resolves tokens in each email, (4) snapshots the store, hands the resolved email to the model runner, snapshots again, and grades the email against the diff, (5) tells the controller it was served, (6) at scenario completion, fetches the full cumulative state and grades the scenario as a whole, (7) ticks the clock.

---

## Token resolution

The Excel sheet uses `{...}` tokens as placeholders for dates and links. The token syntax in the source data is messy and inconsistent — different spacing, different orderings, mixed casing. The resolver normalizes the inside of every token (lowercase + strip whitespace) before matching, so all of these resolve to the same value:

```
{date-nextweek} = {date-next week} = {date-next-week} = {nextweek-date}
```

### Supported tokens

| Token (any spacing) | Resolves to |
|---|---|
| `{date}` | sim date |
| `{date+N}` | sim date + N days |
| `{date+Nweeks}` | sim date + N weeks |
| `{date-tomorrow}`, `{date-nextday}` | sim date + 1 day |
| `{date-thisweek}` (any spacing) | sim date (treated as today, since "this week" is ambiguous) |
| `{date-nextweek}` (any spacing) | sim date + 1 week |
| `{date-next week -N}` | sim date + 1 week − N days |
| `{nextweek-date}` (any spacing) | sim date + 1 week |
| `{nextweek-date +N}` | sim date + 1 week + N days |
| `{nextweek-date -N}` | sim date + 1 week − N days |
| `{date-beginningmonth}` | first day of current month |
| `{date-nextmonth}` | same day-of-month, next month (clamped to last valid day) |
| `{date-nextmonth+N}` | next month, +N days |
| `{date-Nth}` (e.g. `{date-14th}`) | next future occurrence of that day-of-month |
| `{date-10AM}`, `{date-2PM}`, `{date-1:14PM}`, `{date-12:30PM}` | sim date at that time |
| `{link}`, `{meeting-link}` | `[meeting link]` |

### Tokens left as-is (intentionally)

| Token | Why |
|---|---|
| `{Annual Report 1}`, `{contract 2}`, `{Q3 onboarding strategy doc}` | Document references — not date-related, no resolution needed |
| `{date-12:30-2:00PM}` | Time range — ambiguous, defer to team decision |
| `{Tuesday- this week at 3:00 PM}` | Mixed day+time reference — no canonical pattern |
| Anything else unrecognized | Visible in output for review |

The resolver never crashes on an unknown token; it just leaves it untouched.

### Date format

Dates render as `"March 15, 2000"`. Times render as `"March 15, 2000 at 10:00 AM"`.

---

## API

### Public functions

| Function | Purpose |
|---|---|
| `resolve_tokens(text, sim_date)` | Replace every `{...}` in `text` using `sim_date` as the reference. Unknown tokens are left untouched. |
| `apply_date_substitutions(email, sim_date)` | Returns a deep copy of the email with `subject` and `body` token-resolved. |
| `model_interaction_mock(emails, sim_date, verbose=True)` | Fallback mock when `model_runner.py` is not present. Sleeps ~0.05s. |
| `run_simulation(path, sim_days, sim_start, seed, verbose, model_fn, scenarios)` | The full simulation. `model_fn` overrides the model call (used by tests). `scenarios` bypasses Excel loading. Returns `{total_score, total_max, daily_log, remaining_inactive, remaining_active}`. |

### Constants

```python
SIM_START = datetime(2000, 1, 1, tzinfo=timezone.utc)
SIM_DAYS  = 100
```

---

## Connection to the grader

`define_grading_system(input, calendar, todo)` accepts either an `Email` or a `Scenario`. The engine calls it **both ways** during a run, producing two parallel scoreboards:

1. **Per-email grade** — called after every email handoff whose `success_criteria` is set, with `grade_by_scenario=False`. The calendar/todo state passed in is a **diff** of what the AI added or modified during that one email's turn. See "Per-email grading via state diff" below.
2. **Per-scenario grade** — called at scenario completion with the full cumulative state, the loader's aggregated `success_criteria` list, and `grade_by_scenario=True`. Unchanged from prior behavior.

The two grades can disagree, and the gap is informative:
- High email score, low scenario score → the AI handled each email correctly in isolation but the cumulative end state isn't right (often: did the same right thing twice, leaving a duplicate).
- Low email score, high scenario score → individual emails missed their criteria but the final state still matches the scenario's aggregate criteria (often: actions taken in the wrong order, with later emails fixing earlier mistakes).

### Per-email grading via state diff

Naively grading each email against the store's current state would let actions taken for **earlier** emails in the chain falsely satisfy or falsely violate **this** email's criteria — the store is cumulative within a scenario, so it carries forward.

The engine sidesteps this by snapshotting the store before and after each `run_model_turn`, then computing a diff:

```python
state_before = fetch_scenario_results(scenario)
run_model_turn(resolved, sim_date, scenario_id=...)
state_after = fetch_scenario_results(scenario)
email_result = _grade_email_against_diff(resolved, state_before, state_after)
```

The diff classifies each todo and event as **added**, **removed**, or **modified** (same id, different content) by comparing the two snapshots. A synthetic `CalendarResponse` + todos list built from added + modified items is what gets passed to the grader, so:

- `TC-x` and `CC-x` criteria match only items the AI created or updated during this turn.
- "No action" passes only when all six diff lists are empty.
- Deletes don't appear in the synthetic state, so the engine adds an override: if any of the six lists is non-empty, "No action" sub-criteria are flipped from pass to fail (`_diff_took_action`).

Helper functions: `_state_diff`, `_diff_took_action`, `_grade_email_against_diff` (all in `engine.py`).

### Scenario filtering

`pipeline.fetch_scenario_results()` converts the loader's string `scenario_id` (e.g. `"T01"`) into the hashed int that the store uses, then filters `store.todos_db` and all calendar events to only those matching that int. The grader itself doesn't need to filter — it receives pre-scoped data.

State accumulates across scenarios and is never cleared mid-run. This is intentional — the benchmark tests how the model handles a growing store.

### When the grader is upgraded for exact-date checking

Use `from engine import resolve_tokens` and call it on each criteria string before parsing. Same resolver as the engine uses — the dates will line up.

---

## What `run_simulation` returns

```python
{
    "total_score": int,        # scenario-level total (1 point per scenario with all criteria met)
    "total_max": int,
    "email_score": int,        # per-email total (1 point per email criterion met, diff-based)
    "email_max": int,
    "daily_log": [
        {"day": 1, "date": "2000-01-01", "served": 3, "score": 0, "max_score": 1},
        ...
    ],
    "by_type": {               # scenario-level score grouped by scenario_type
        "T": {"count": 18, "score": 12, "max_score": 18},
        ...
    },
    "by_type_email": {         # per-email score grouped by scenario_type
        "T": {"count": 47, "score": 31, "max_score": 47},
        ...
    },
    "remaining_inactive": int, # scenarios never activated (should be 0 for normal runs)
    "remaining_active": int,   # chains overflowing past day 100 (rare, see Flow Controller doc)
}
```

---

## Model runner integration

The engine tries to import `run_model_turn(email, sim_date)` from `model_runner.py`. If the module isn't present, it falls back to `model_interaction_mock` (prints and sleeps). The `model_fn` parameter on `run_simulation()` overrides both — used by the E2E tests to inject a stub.

Priority: `model_fn` > `run_model_turn` (from model_runner) > `model_interaction_mock`.

---

## E2E testing

`tests/test_e2e.py` runs the full 100-day simulation with hand-crafted scenarios and a stub model that writes directly to the in-memory store. Two tests:

- **`test_perfect_stub_scores_max`** — stub always creates the right todo or event, asserts `total_score == total_max`
- **`test_bad_stub_scores_only_no_action`** — stub does nothing, asserts only the "No action" scenario passes (1/3)

Run with `python -m pytest tests/test_e2e.py -v`.

---

## Notes & gotchas

- **Reproducibility**: seeded with `seed=42` by default. Same seed → same scenario distribution, same chain offsets. Pass `seed=None` for nondeterministic runs.
- **Sim start date**: `January 1, 2000` is arbitrary, chosen to match the calendar model's example. Change it via the `sim_start` parameter — token resolution adapts automatically.
- **Verbose mode**: `verbose=True` prints day-by-day activity. Turn off for batch runs.
- **Chain order**: email N in a scenario can't be served until emails 1..N-1 are served. Enforced by `flow_controller.py`.
- **Day 100 cap**: chain offsets are clamped so no scenario spills past the last day.
- **State accumulates**: todos and calendar events are never cleared between scenarios. The grader scopes by `scenario_id` via `pipeline.fetch_scenario_results()`.

---

## What's intentionally not implemented

- **Resolving criteria tokens** — left for when the grader is upgraded to exact-date checking
- **Persistent results** — `run_simulation` returns a dict; saving to JSON/CSV is the caller's job
- **Mid-run state inspection** — no debug hooks beyond `controller.status()`. Add as needed.
