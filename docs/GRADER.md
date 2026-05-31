# grader.py — Systems Grader

`grader.py` exposes a single function, `define_grading_system(input, calendar, todo)`, which takes either an `Email` or `Scenario` object along with a `CalendarResponse` (from `app.models.calendar`) and a `list[TodoResponse]` (from `app.models.todo`) and returns a score dict. It first checks whether the input is a single email or a full scenario, if it's an email, it wraps that email's `success_criteria` string into a list, and if it's a scenario, it uses the already-collected `success_criteria` list directly. Each criteria string is then checked against the tool state using prefix-level matching via the `_check_criteria` helper. **This will need to be updated once we decide what the final success criteria looks like.**

## Scenario-Scoped Grading

The grader itself does not filter by `scenario_id` — that happens upstream. `pipeline.fetch_scenario_results(scenario)` pulls only the todos and calendar events that belong to the scenario being graded (matched by hashed `scenario_id`), then the engine passes that filtered state to the grader. This means the grader always sees only the data relevant to the current scenario, even though the store accumulates state across all scenarios during a run.

## Prefix-Level Checking

The grader parses the prefix of each criteria string to determine what type of action the model should have taken, then verifies that action against the current calendar and todo state:

| Prefix | Expected Action | Check |
|---|---|---|
| `TC-` | Model should have created a todo | `len(todo) > 0` |
| `CC-` | Model should have created a calendar event | `len(calendar.events) > 0` |
| `RS-` | Model should have rescheduled an event | `len(calendar.events) > 0` |
| `No action` | Model should have done nothing | `len(todo) == 0` AND `len(calendar.events) == 0` |

Criteria strings without a recognized prefix (freeform like "Add task", "Delegate Task") are **not graded** — see "Free-text criteria" below (FIX-5).

A single criteria string can contain multiple prefixes (e.g. `TC-{date} && CC-{date}`). The current implementation checks all matching prefixes — if any expected action is missing, the criteria fails.

## Free-text criteria are ungraded, not auto-passed (FIX-5, Sprint 5)

A criterion that is *entirely* free-text (no `TC`/`CC`/`RS` prefix, not "No action") is **excluded** from `score` and `max_score` and reported in a new `ungraded` list on the result dict, instead of silently auto-passing and inflating the max. So the score reflects only what the grader can verify. A mixed criterion (`TC-{date}, Flag`) is still graded on its real sub-check; the free-text sub adds no pass/fail.

`define_grading_system(...)` returns `{"score", "max_score", "details", "ungraded"}`.

### Splitter repairs

`_parse_sub_criteria` used to fragment real prefixed criteria into auto-passing junk (measured on `Emails.xlsx`; see SPRINT5_REMEDIATION Appendix C). Fixed:
- **Stray `^`** — a lone `^` is a dataset placeholder, dropped (was 13 phantom entries).
- **`or` alternatives** — `CC-{A}, or CC-{B}, or CC-{C}` is one alternative group, merged into a single CC sub that passes if *any* branch's resolved date matches.
- **Leading labels** — `Update:  TC-{...}` has its label stripped so the `TC` prefix is recognized.

Net effect: free-text sub-criteria dropped from 30 to 12, and those 12 are now reported as `ungraded` rather than silently passing.

### Known follow-up (not in scope)

Three remaining free-text criteria describe real actions without a prefix (`delete meeting {date-14th} {date-11AM}`, `Remove meeting on {date-1:14PM}`, `create new meeting on {date-3PM}`). They are reported as `ungraded`. The honest fix is to normalize them in `Emails.xlsx` (rewrite as `RS-`/`CC-`) or teach the grader the verbs — better than the old auto-pass, but worth recovering later.

## Date and content verification (FIX-2)

Date verification **is** implemented. The engine resolves date tokens in `success_criteria` before grading, keeping the braces: `CC-{date}` served on March 15, 2000 becomes `CC-{March 15, 2000}`. `_extract_date_token` reads the resolved date from inside the braces and `_event_matches_date` compares it to the event's start, so an event on the wrong date fails instead of passing a count-only check. Content tokens (`TC-item3`) are matched against todo title/description. Tokens that stay unresolved (e.g. `{nextweek-wednesday}`, which the resolver doesn't handle) fall back to the count check.

**Time-of-day tokens** (`CC-{date-3PM}`) resolve to a date+time string (`March 15, 2000 at 03:00 PM`); the grader matches date **and** hour/minute, so an event at the wrong time fails too. `_event_matches_date` is **fail-closed**: a target date string it cannot parse is a non-match, not a free pass. To avoid over-failing, `_cc_sub_passes` only runs a date check for branches whose extracted token is a *parseable* date — a braced non-date placeholder (e.g. the stray `CC-{C}` in the dataset) is treated as "no date" and falls back to the count check. (Before this, an unparseable time-of-day target hit a `return True` fallback that let a wrong-date event satisfy a timed criterion — found by the Sprint-5 adversarial gate.)
