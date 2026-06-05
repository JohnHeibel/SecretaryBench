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

## 2. Global calendar + global grading

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
