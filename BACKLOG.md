# Backlog (deferred features, decided "not now")

Wanted-someday features we deliberately parked, with the reason. Different from
`OPEN_QUESTIONS.md` (design calls still being decided) — these are decisions to *defer*, not
open questions. Revisit only when the trigger below is real.

## 1. Ranged / list offsets — per-seed "random day" (anti-memorization)

**What:** an offset written as a set or range, e.g. `@kickoff+[2,5,9]w` or `+[2..14]d`. The run's
seed picks one value, so the same scenario lands on a **different day each run** — a hedge against
memorization / training-set contamination.

**It can be deterministic** (verified in design): pick the value ONCE per seed at plan time, freeze
it into the spine/anchor, and have both the rendered email and the grader read that one frozen value
(single source of truth). Re-running a seed stays byte-identical; within a run there is still exactly
one correct date, and it is the one the model read in the email.

**Why it's parked:**
- **Low discrimination.** It is a synthetic knob; the Haiku runs scored ~85% regardless of synthetic
  difficulty / span / bait. The email has to reveal the picked number, so it does not make the
  reasoning harder — it only resists answering *without reading*.
- **It costs determinism guarantees.** It breaks the current seed-invariant-dates property and makes
  **feasibility seed-dependent**: a scenario that is oracle-1.0 on one seed can be infeasible on
  another (a large pick can collapse a reschedule window). So the "oracle solves 1.0" safety net would
  have to be checked across the whole range × seeds, not a single run — real authoring/validation cost.
- **Serve-date variation already prevents absolute-date memorization** today (the answer's absolute
  date already changes per run).

**Cheaper interim if variety is ever wanted:** generate fixed-offset template *variants* at build time
(`sb/scale.py`) — same spread of days, each a normal deterministic, independently-feasible node. Keep
the variety in corpus construction, out of the grammar/runtime.

**Trigger to revisit:** the benchmark is published and training-set contamination becomes a real,
measured concern.

**Build sketch (when picked up):** parse `[a,b,c]` / `[a..b]` in the offset (`sb/resolver.py` + the
`webapp/lib/dateExpr.ts` mirror) → pick once in `build_plan` via the existing seeded RNG, keyed by a
stable anchor id → freeze into the spine/anchor table → tests (same seed reproduces, seeds vary,
body == grader every time) → webapp UI to enter the list → update the determinism-contract wording in
`GRADING_MODEL.md` (the seed would then also pick ranged dates).

## 2. Global calendar + global grading (includes cross-storyline time conflicts)

**What:** today grading is **scoped per scenario** (attribution-scoped) — two scenarios can both book
Tue 2 PM and each passes on its own. A future, harder mode would grade against a single **global**
calendar: global coherence, no cross-scenario double-booking, possibly "find an open slot."

**Why scoped now:** scoped grading is what keeps everything **deterministic and attributable** — a
wrong action can only ever cost its own email's point, never another scenario's score. Global grading
reintroduces non-determinism: which scenario "owns" a contested slot has no unique answer, and
find-a-free-slot has no single correct date (exactly why free-slot / `no_overlap` was cut). Global
double-booking is already tracked **monitor-only** (never scored) in the current model.

**Trigger to revisit:** someone wants to take on a harder, global-coherence variant of the benchmark.

**If someone takes it on:** they need a *deterministic* global-coherence rule, or keep the global
metrics monitor-only (unscored) while scoped grading stays the scored path. Keep scoped grading as the
default either way — it is the property that makes the score clean.

### 2a. Cross-storyline time conflicts (the "double-booking" sub-problem)

An exec-meeting ask (2026-06-05) was to deal with two storylines' events colliding on the same clock
slot — e.g. both book Tue 2–3 PM. **This lives here, not as a near-term fix, because under scoped
grading a clash is purely cosmetic: it can never change a score.** It only becomes a real problem the
day we grade against one global calendar (this epic). So we parked it here rather than complicate the
serving algorithm for a non-scored artifact.

**Measured, per the "measure before deciding" request.** The monitor is built and shipped:
- `python -m sb.conflicts --corpus <dir>` — resolves every served event to its FINAL clock slot
  (replaying create/move/cancel per obligation), then counts cross-storyline `TimeInterval` overlaps.
  Same-storyline overlaps and ambiguous predicates (`any_of` / `by` / bare-day) are excluded; it also
  reports peak concurrency. This *is* the "monitor-only" tracking the section above refers to.
- `python -m sb.probe` — synthesizes a ~20-author stand-in corpus (the real corpus lives in the prod
  store; `corpus/nodes/` is empty in git), so the monitor has something to run on offline.
- `sb/tests/test_conflicts.py` — locks the replay (move supersedes, cancel vacates) + overlap rules.

**Finding (probe, 20 authors × 3 timed events, 5 seeds each):** conflicts are **common, not rare** —
**18%–49% of timed events** collide depending on how clustered the meeting times are (worst: everyone
picks the same handful of times with no week-spread; best: varied times spread over 8 weeks). Peak
concurrency stayed ~3. **Every regime still oracle-solves 1.0**, confirming conflicts never touch the
score. The exec's two branches were *(a) resample the serve order when conflicts are rare* vs *(b) add
non-overlap as a serving constraint when common*; the probe points at (b) — but the **decision must use
the real number**, not the probe.

**When this epic is picked up:**
1. Pull the real corpus (`python -m sb.sync`) and run `python -m sb.conflicts --corpus corpus` for the
   true rate.
2. If still common and we want global coherence, add non-overlap as a *deterministic* serving
   constraint (an event can't be placed on a slot another storyline already holds) — but note this
   couples storylines that are otherwise independent, so it has to stay compatible with the
   reproducible-plan guarantee.
3. Keep scoped grading as the scored path regardless; the global calendar stays monitor-only unless a
   deterministic global rule is found (see the parent section).
