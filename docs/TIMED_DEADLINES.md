# Next step: timed `by` deadlines, drop the scaffold button, slim Project Atlas

Status: **active, next up.** Not backlog. All work happens on the `webapp-v2` branch (do not switch,
create, or merge branches; do not push). Three parts, do them in order — Part 3 (the Atlas reshape) is
last because it should showcase the Part 1 feature and assumes the Part 2 button is already gone.

---

## Part 1 — a clock time on the "on or before" (`by`) deadline

### The ask in one line

In the answer-key builder, "on exactly" (`eq`) and "any of" (`any_of`) let an author pick a clock time;
"on or before" (`by`) does not. Authors who tested the tool were confused by that gap — real deadlines
have times ("get the budget to me **by Friday 5 PM**"). Make a time on `by` real and gradeable.

### Why it's worth doing

- Realistic and idiomatic — clock-time deadlines are how people talk. The current day-only `by` is one of
  the few places the benchmark is deliberately fuzzy; a timed deadline tightens it.
- It fits the grading idiom: still exact, still deterministic, still a pure function of (answer key, serve
  date, ancestors). A timed `by` is just an upper-bound datetime instead of an upper-bound day.

### The semantics to implement

> `by <datetime>` = "the created object **starts at or before** that datetime, and not before the serve day."

- A **bare-date** `by` keeps today's behavior exactly: "any day up to and including the deadline day."
- A **timed** `by` (the expr resolves to a `datetime` or `TimeInterval`) compares at clock granularity: an
  object on an earlier day passes at any time (Thu 11 PM ≤ Fri 5 PM); on the deadline day it must land at
  or before the cutoff. The cutoff is the **start** of a `TimeInterval` (`@HH:MM-HH:MM`) or the `datetime`
  itself (`@HH:MM`). **Start-only** — a deadline is about when the thing lands, not its length.

### Why it's a SMALL change: 3 of the 4 engine pieces already cope

- **Oracle — no change.** `_target` resolves the `by` expr verbatim (`sb/oracle.py:26-27`) and `_placement`
  already places a `datetime`/`TimeInterval` at its exact time (`sb/oracle.py:54-64`). So `by next:fri
  @17:00` already gets placed at Fri 5 PM, which is ≤ the cutoff → oracle-solves with zero edits. Add a
  confirming test, but expect no code change here.
- **Scheduler — no change.** Serve-by windows are day-level by design (`deadlines: dict[str, date]`
  `sb/scheduler.py:45`; `min(dates)` `sb/scheduler.py:99`). Coercing a timed deadline to its date for
  *serving* is correct — the time only matters for *grading the object*, never for which day the email arrives.
- **Grader — the one real change** (below).

### Plan, file by file

1. **`sb/grader.py`, the `by` branch of `_predicate_ok` (currently lines 110-111).** Today:
   ```python
   if "by" in predicate:
       return ctx.serve <= obj.when.date() <= _to_date(resolver.resolve(predicate["by"], ctx))
   ```
   Branch on the resolved type. If it's a `datetime` or `TimeInterval`, compare `obj.when` (a datetime) to
   the cutoff and floor the lower bound at the serve day's start; otherwise keep the day-level path. Sketch:
   ```python
   if "by" in predicate:
       dl = resolver.resolve(predicate["by"], ctx)
       if isinstance(dl, (datetime, TimeInterval)):
           cutoff = dl.start if isinstance(dl, TimeInterval) else dl
           return datetime.combine(ctx.serve, time.min) <= obj.when <= cutoff
       return ctx.serve <= obj.when.date() <= _to_date(dl)
   ```
   (`_matches_value` at `sb/grader.py:83-99` is the reference for how timed values are compared — mirror its
   type handling. Import `time`/`TimeInterval` as needed.)
2. **`sb/grader.py`, `_wrong_when_reason` (around line 166).** Let a boundary-day near-miss explain itself
   ("landed Fri 6 PM, deadline was Fri 5 PM"). `_describe_predicate`/`_fmt_value` (lines 149-163) already
   render datetimes with the time, so the "by Fri 5 PM" text comes for free.
3. **`webapp/components/AnswerKeyBuilder.tsx`** — flip two gates:
   - ~line 206, `allowTime={ek === "event" && po === "eq"}` → also allow it for `po === "by"` on events.
   - ~line 189, the `dropTime(...)` call that strips the clock when switching the predicate to `by` — stop
     stripping for `by` so a time set there survives.
   The clock UI already exists in `DateBuilder.tsx` (`TimeControl`); it's purely gated off today.
4. **Re-vendor**: `python3 webapp/scripts/vendor_sb.py`, then `--check` must stay green.
5. **Tests** (`sb/tests/`): grader cases — pass before the cutoff on the deadline day, fail after, pass at
   any time on an earlier day, bare-date `by` unchanged. Plus one oracle round-trip with a timed `by`
   (build_corpus → build_plan → engine.run(oracle_model); assert score == 1.0).
6. **Docs**: `GRADING_MODEL.md` and `ANSWER_KEY_GRAMMAR.md` frame `by` as "the day-level deadline." Update to
   "day-level deadline, or a datetime cutoff if you give it a time." Update the `by`/time line in the root
   `CLAUDE.md` grammar/gotchas section to match.

### Decisions to lock (with recommended defaults)

1. **Start-only vs start+end for events.** `eq` checks both (`sb/grader.py:90`). For a deadline compare the
   **start** only — matches the oracle and the natural reading. **Default: start-only.**
2. **This deliberately changes engine semantics.** It edits `sb/grader.py`, crossing the standing rule "make
   the UI fit the engine, don't change the engine." Fine because we're choosing to — but run the **full**
   `pytest sb/tests` and re-confirm the corpus still oracle-solves 1.0, not just a re-vendor.
3. **To-do deadlines: time or events-only?** Events clearly should. A timed to-do deadline is coherent but
   check what the live harness defaults a no-time to-do's clock to (`sb/live/`) — on the boundary day that
   default decides pass/fail. **Default: enable for events first; decide to-dos after checking the default.**

---

## Part 2 — delete the "scaffold an action for each date" button

It's a typing shortcut nobody needs; remove it.

- **`webapp/components/AnswerKeyBuilder.tsx`** — delete the scaffold logic and its button: `usedAnchors`,
  `scaffoldable`, `isBlankOp`, and `scaffoldFromBody` (the block around lines 111-121) and the button JSX
  (around lines 253-258). **Keep the `bodyAnchors` prop** — it's still used at line ~221 (`crossRefs =
  anchorRefsIn(pred).filter((n) => !bodyAnchors.includes(n))`), so only remove the scaffold-specific code,
  not the prop or the `ownBodyAnchors` wiring in `EmailEditor.tsx`.
- **`webapp/app/guide/page.tsx`** — remove the whole "An email with lots to do at once" Card (lines ~98-99);
  it only existed to explain this button. (Drop the card; the multi-action workflow is just "add another
  action" now.)
- **Tooltips that advertise it** — the Atlas tour blurb in `Sidebar.tsx` (~line 37) lists "many actions in
  one email" as a feature; drop that phrase (and it falls out naturally in Part 3 anyway).
- No engine or test impact — this is pure UI.

---

## Part 3 — slim Project Atlas to ~8 emails (do LAST)

Goal: Atlas is the single loadable example and the basis for a how-to video. Today it's 18 emails — too long
to narrate. Cut it to **~8 emails that are small but still rich**: a tight, coherent product-launch story
that tours the headline constructs, **including the new timed `by` deadline from Part 1.** Lives in
`webapp/lib/templates.ts` (`projectAtlasNode`).

### Recommended 8-email lineup (tune freely; the ★ ones are must-keep)

1. **freeze** ★ — EVENT with a clock time; body publishes `@atlas_freeze` + `@atlas_launch`. [event + time + the two load-bearing anchors]
2. **ceo-note** — CEO-sent (`from: CEO`) instruction; a needle off `@atlas_launch`. [CEO-sent + needle]
3. **legal-fyi** ★ — NO action, a distractor that looks like a task. [restraint — a core concept]
4. **beta** ★ — a TO-DO due BY a **timed** deadline (e.g. business days after freeze, by 5 PM) = a needle. [todo + **timed by** + needle — showcases Part 1]
5. **board-demo** ★ — EVENT, a needle two weeks after freeze. [needle, medium span]
6. **demo-moved** ★ — RESCHEDULE: move the board demo. [move]
7. **launch-dinner** — EVENT on launch night, a needle reusing `@atlas_launch` (the longest span). [needle, longest span; sets up the cancel]
8. **dinner-cancel** ★ — CANCEL that launch dinner. [cancel]

That tours: event+time, multi-anchor setup, CEO-sent, needles at several spans, no-action restraint, a
**timed deadline**, reschedule, and cancel. `any_of` is dropped to hit 8 — if you'd rather show it, swap it
in for the launch-dinner/cancel pair (but then you lose cancel). Keep the cast on the standard roster and
the sender-only model (every email `to: CEO`, no To/Cc).

### Hard constraints

- **Atlas must still lint clean AND oracle-solve 1.0**, standalone and at scale. Verify with
  `npx tsx webapp/scripts/checkTemplate.mts` after reshaping — that's the gate the current template passes.
- Rewrite the big numbered tour comment at the top of `templates.ts` to match the new lineup.
- Update the Atlas-tour copy in `Sidebar.tsx` (~line 37 tooltip) and `Workspace.tsx` (`StartEmpty`, ~line
  327) so the "tours every feature" list matches the 8 emails actually present.

---

## Verify before declaring done (run all)

```bash
pytest sb/tests                                   # full suite (Part 1 changes engine semantics)
python -m sb.demo                                 # oracle round-trip still 1.0
cd webapp && npm run typecheck
python3 webapp/scripts/vendor_sb.py --check       # anti-drift gate stays green
npx tsx webapp/scripts/dateExpr.test.mts          # date grammar round-trips vs the real resolver
npx tsx webapp/scripts/checkTemplate.mts          # Project Atlas lints + oracle-solves 1.0
```

## Guardrails

- **Stay on `webapp-v2`.** Do not switch, create, or merge branches; do not push.
- Part 1 touches only the `by` predicate's grading — leave `eq`/`any_of`/`in`/`not_in`, the scheduler's
  serve algorithm, and anything about cross-storyline conflicts (parked epic, BACKLOG §2a) alone.
- Commit/PR style: no em-dashes, no Claude/co-author attribution, conversational tone.
