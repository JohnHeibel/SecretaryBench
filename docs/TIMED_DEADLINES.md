# Next step: a clock time on the "on or before" (`by`) deadline

Status: **active, next up.** Not backlog. Work happens on the `webapp-v2` branch (do not switch or create branches).

## The ask in one line

In the answer-key builder, "on exactly" (`eq`) and "any of" (`any_of`) let an author pick a clock time;
"on or before" (`by`) does not. Authors who tested the tool were confused by that gap — real deadlines
have times ("get the budget to me **by Friday 5 PM**"). Make a time on `by` real and gradeable.

## Why this is worth doing

- It's realistic and idiomatic. Clock-time deadlines are how people actually talk. The current day-only
  `by` is one of the few places the benchmark is deliberately fuzzy; a timed deadline tightens it.
- It fits the grading idiom: still exact, still deterministic, still a pure function of (answer key,
  serve date, ancestors). A timed `by` is just an upper-bound datetime instead of an upper-bound day.

## The semantics to implement

> `by <datetime>` = "the created object **starts at or before** that datetime, and not before the serve day."

- A **bare-date** `by` keeps today's behavior exactly: "any day up to and including the deadline day."
- A **timed** `by` (the expr resolves to a `datetime` or `TimeInterval`) compares at clock granularity:
  an object on an earlier day passes at any time (Thu 11 PM ≤ Fri 5 PM); on the deadline day it must land
  at or before the cutoff. The cutoff is the **start** of a `TimeInterval` (`@HH:MM-HH:MM`) or the
  `datetime` itself (`@HH:MM`). **Start-only** — a deadline is about when the thing lands, not its length.

## Why it's a SMALL change: 3 of the 4 engine pieces already cope

- **Oracle — no change.** `_target` resolves the `by` expr verbatim (`sb/oracle.py:26-27`) and
  `_placement` already places a `datetime`/`TimeInterval` at its exact time (`sb/oracle.py:54-64`). So a
  `by next:fri @17:00` deadline already gets placed at Fri 5 PM, which is ≤ the cutoff → it oracle-solves
  with zero edits. (Add a confirming test, but expect no code change here.)
- **Scheduler — no change.** Serve-by windows are day-level by design (`deadlines: dict[str, date]`
  `sb/scheduler.py:45`; `min(dates)` `sb/scheduler.py:99`). Coercing a timed deadline to its date for
  *serving* is correct — the time only matters for *grading the object*, never for which day the email
  arrives.
- **Grader — the one real change** (below).

## Plan, file by file

1. **`sb/grader.py`, the `by` branch of `_predicate_ok` (currently lines 110-111).** Today:
   ```python
   if "by" in predicate:
       return ctx.serve <= obj.when.date() <= _to_date(resolver.resolve(predicate["by"], ctx))
   ```
   Branch on the resolved type. If it's a `datetime` or `TimeInterval`, compare `obj.when` (a datetime) to
   the cutoff and floor the lower bound at the serve day's start; otherwise keep the existing day-level
   path unchanged. Sketch:
   ```python
   if "by" in predicate:
       dl = resolver.resolve(predicate["by"], ctx)
       if isinstance(dl, (datetime, TimeInterval)):
           cutoff = dl.start if isinstance(dl, TimeInterval) else dl
           return datetime.combine(ctx.serve, time.min) <= obj.when <= cutoff
       return ctx.serve <= obj.when.date() <= _to_date(dl)
   ```
   (`_matches_value` at `sb/grader.py:83-99` is the existing reference for how timed values are compared —
   mirror its type handling. Import `time`/`TimeInterval` as needed.)
2. **`sb/grader.py`, `_wrong_when_reason` (around line 166).** Let a boundary-day near-miss explain itself
   ("landed Fri 6 PM, deadline was Fri 5 PM"). `_describe_predicate`/`_fmt_value` (lines 149-163) already
   render datetimes with the time, so the "by Fri 5 PM" text comes for free.
3. **`webapp/components/AnswerKeyBuilder.tsx`** — flip two gates:
   - line ~206, `allowTime={ek === "event" && po === "eq"}` → also allow it for `po === "by"` on events.
     (Decide whether to-do deadlines get a time too — see decision 3.)
   - line ~189, the `dropTime(...)` call that strips the clock when switching the predicate to `by` —
     stop stripping for `by` so a time set there survives.
   The clock UI already exists in `DateBuilder.tsx` (`TimeControl`); it's purely gated off today.
4. **Re-vendor**: `python3 webapp/scripts/vendor_sb.py` (keeps `lib/schema.generated.ts` and the vendored
   copy in sync), then `python3 webapp/scripts/vendor_sb.py --check` must stay green.
5. **Tests** (`sb/tests/`): add grader cases — pass before the cutoff on the deadline day, fail after it,
   pass at any time on an earlier day, and the bare-date `by` still behaves as before. Add one oracle
   round-trip with a timed `by` (build_corpus → build_plan → engine.run(oracle_model); assert score == 1.0).
6. **Docs**: `GRADING_MODEL.md` and `ANSWER_KEY_GRAMMAR.md` frame `by` as "the day-level deadline." Update
   to "day-level deadline, or a datetime cutoff if you give it a time." Update the `by`/time line in
   the root `CLAUDE.md` grammar/gotchas section to match. (The existing line "bare dates grade day-level;
   timed exprs grade to the minute" already half-covers this.)

## Decisions to lock before/while coding (with recommended defaults)

1. **Start-only vs start+end for events.** `eq` checks both start and end (`sb/grader.py:90`). For a
   deadline, compare the **start** only — matches what the oracle produces and the natural reading.
   **Default: start-only.**
2. **This deliberately changes engine semantics.** It edits `sb/grader.py`, which crosses the standing
   working agreement "make the authoring UI fit the engine, don't change the engine." That's fine *because
   we're choosing to*, but it means: run the **full** `pytest sb/tests` and re-confirm the existing corpus
   still oracle-solves 1.0, not just a quiet re-vendor. Treat it as an engine change, not a UI tweak.
3. **Do to-do deadlines get a time, or only events?** Events clearly should. A timed to-do deadline is
   coherent too, but check what the live harness defaults a no-time to-do's clock to (`sb/live/`) — on the
   boundary day that default now decides pass/fail. A model can always dodge it by landing *before* the
   deadline day. **Default: enable for events first; decide to-dos after checking the live default.**

## Verify before declaring done

```bash
pytest sb/tests                              # full suite, not just the new cases
python -m sb.demo                            # oracle round-trip still 1.0
cd webapp && npm run typecheck
python3 webapp/scripts/vendor_sb.py --check  # anti-drift gate stays green
npx tsx webapp/scripts/dateExpr.test.mts     # date grammar round-trips vs the real resolver
```

## Guardrails for whoever picks this up

- **Stay on `webapp-v2`.** Do not switch, create, or merge branches. Commit here.
- Touch only the `by` predicate's grading. Do not change `eq`/`any_of`/`in`/`not_in`, the scheduler's
  serve algorithm, or anything about cross-storyline conflicts (that's a separate parked epic, BACKLOG §2a).
- Commit/PR style: no em-dashes, no Claude/co-author attribution, conversational tone.
