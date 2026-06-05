# Storyline Grading Model

How SecretaryBench grades a long-horizon temporal-reasoning run, and why it stays
deterministic while the calendar looks realistically messy. This is the contract;
`HOW_IT_WORKS.md` is the plain-language version, `ANSWER_KEY_GRAMMAR.md` the
token/answer language, `SERVING_AND_SCHEMA.md` the DAG/serve mechanics.

## Goal

Measure **long-horizon temporal reasoning**: can an AI assistant track many
time-dependent threads across a long simulated span, resolve relative dates
("two weeks after the kickoff"), and keep a coherent calendar — including
rescheduling and resolving authored conflicts.

The benchmark's job is a **clean, attributable, reproducible** number. Any failure
that isn't *about temporal reasoning* is contamination and must be designed out.

## Two layers (one is model-proof)

- **Spine (immutable):** serve plan, anchor table, every expected date, deadlines,
  feasibility. All computed by the scheduler *before any model runs*, straight from
  the corpus. The model has no tool that writes here. A child that resolves a date
  relative to a parent reads the precomputed anchor, never a calendar object.
- **Workspace (mutable):** the events/todos the model creates in the store. The only
  thing the model can change.
- **Memory (immutable, the model's ground truth):** the **inbox** — append-only,
  read-only to the model. Every instruction and fact is recoverable via
  `search_inbox`. The calendar is *reconstructible from the inbox*, never the source
  of truth.

The grader takes **expected dates from the spine** and **observed objects from the
workspace**. So corrupting the workspace can never change a correct answer, and a
competent model can always recover lost workspace state from the inbox.

## Scenario = node = storyline

A scenario is one node: a coherent thread (shared cast, named obligations, a DAG of
date dependencies) over many simulated days.

- **Authored & verified in isolation** — each scenario must be oracle-solvable
  *alone* (`build_corpus([node])` → score 1.0). That's the authoring safety net.
- **Executed together** — at run time many scenarios share one inbox and one
  calendar, interleaved by the random batch. That shared, busy world is what creates
  the long-horizon challenge (concurrency, retrieval under noise, long span).

## Determinism = seed-invariant per-email grading

Each email's grade is a **pure function of (its answer key, its serve date, its
deterministic DAG ancestors)**. The random seed only shuffles *which* scenarios run,
*when* emails arrive, and *how much* noise surrounds them — never the within-node
ordering, the resolved dates, or the correct answer. Re-running a seed is
byte-identical; a correct model scores the same on every seed.

A correct answer need not be **unique** — a "due by Friday" to-do passes on any day
from arrival through the deadline; "either Tuesday or Thursday" passes on either. The
grader checks "is this a valid member of the deterministic acceptable set," not "is
this THE answer." The set is still fixed by the answer key, never by the random batch.

Otherwise **matching is exact** — exactly one day and time. The old `within:Nd`
tolerance knob is dropped (default and only setting: exact). The *only* deliberate
multi-day answers are the `by` (deadline) and `any_of` (options) predicates above; an
author picks those on purpose, they are not a fuzzy slider.

## The grading contract (binary, per email, scored at the scene)

Each email scores **1** iff, on its turn, the model:

1. **satisfied its authorized ops** — each `create`/`move`/`cancel` resolves to
   exactly one matching object (by keyword, within the node), on the right
   day/time; a `cancel` leaves none; and
2. **took no unauthorized action** — created or mutated nothing beyond what its ops
   authorize (no junk objects in its own thread; no reach into another scenario).

Otherwise it scores **0**. No partial credit. "No-action" (zero-ops) email is the
special case of rule 2 with an empty authorized set.

**The turn is a day.** Each simulated day the model gets that day's batch of emails and
works them with as many tool calls as it needs. Every object it creates or edits is
tagged with the email it was acting on, so each email is graded at the **end of its
serve day** against the actions attributed to it — a stray action is caught then, not
lazily after some other scenario was already graded.

This is deterministic and fair: the oracle does exactly its authorized ops and
nothing else, so it scores 1 on every seed; the rule depends only on the answer key
plus the model's observed actions.

### The thread fence (block + dock)

Cross-scenario corruption is made structurally impossible at the store, not left to
chance or to careful naming. The model passes the email it is acting on with every
edit/delete; the store already knows which scenario owns each object. If the target
belongs to a *different* scenario, the store:

- **blocks** the edit (returns a clear "not your thread" error) — so the other
  scenario's objects, and its score, are never touched; and
- **docks** the acting email — reaching across threads is an unauthorized action, so
  that email scores 0 (rule 2).

A correct model never trips it: every task is about its own thread's objects, so
there is never a legitimate reason to reach into another scenario. Reads still see
the whole calendar (realism); only *edits* are fenced. The fence is invisible to good
behavior and fires only on a genuine mistake.

## Conflicts are authored, never emergent

We test the CEO's calendar getting crowded — but the model never has to *invent* a
free slot (that would have no single deterministic answer). Instead:

- **Every time is fixed, and conflicts are authored as reschedules.** The story itself
  creates the clash and dictates the fix: day 3 "Board meeting Tue 2–3 PM" → day 12 "An
  investor call must take Tue 2–3 PM; move the board to 4–5 PM." The model faces a real
  conflict and must reschedule, but the correct outcome is a fixed state (board at 4,
  call at 2), so it grades cleanly. Uses the `move`/`cancel` verbs we already have.
- **Cross-scenario overlap on the shared calendar is fine.** Grading is
  attribution-scoped, so two scenarios both at 2 PM each pass independently; the
  overlap is realistic noise. The uniqueness rule + rule 2 mean the model is
  *penalized* for "tidying up" things it wasn't asked about — correct behavior.
- **Global double-booking is monitor-only**, never scored (it can't be deterministic).

There is **no "find an open slot" task and no `no_overlap` predicate** — that was the
only non-deterministic, hard-to-author piece, and it's deliberately cut.

## Distinctive naming (nice-to-have, not required)

Distinct cast/subject per scenario is good practice — it makes the model's job clearer
and reads more realistically — but it is **not load-bearing**, because the thread
fence already guarantees correctness:

- A model that grabs the wrong look-alike *instead of* the right object fails its own
  task anyway (its object isn't where it should be — state-based grading), fence or
  not.
- A model that does its task correctly *and also* reaches across threads has
  demonstrably identified the right object, so the extra reach is gratuitous — fair to
  dock regardless of names.
- The model can always disambiguate from the **inbox** (each object carries the
  `email_id` that created it), so identical titles make a task *harder*, never
  *unfair*.

So we don't enforce it. It's a difficulty/clarity lever, not a correctness
requirement.

## Why it's robust (the scare answered)

A wrong action can't cascade across scenarios:

- It can't reach the spine (anchors/expected dates are precomputed, model-proof).
- It can't reach another scenario's objects — the **thread fence** blocks the edit,
  so no other scenario's state (or score) is ever corrupted.
- It can't reach another scenario's grade — grading is attribution-scoped per node.

The only place a model's mistake lands is **its own score**: the offending email is
docked, and any state it fumbled *within its own thread* it can rebuild on a later
turn using the incorruptible inbox (every fact is recoverable via `search_inbox`). A
model that corrupts its own workspace and fails to recover will fail later tasks of
its own — fair and intended, since maintaining coherent state over a long horizon is
the skill under test.

## Status

- **Phase 1 (done):** time-of-day / interval grammar (`@HH:MM[-HH:MM]`, work hours
  05:00–23:00), minute-granular grading of fixed-time events, oracle places at the
  resolved time. (`sb/resolver.py`, `sb/grader.py`, `sb/oracle.py` + tests.)
- **Phase 2 (next, decided):** the **thread fence** (store gates edits/deletes by
  scenario: block + dock) — model passes its acting email on mutations, store checks
  ownership, grader docks the email; generalize rule 2 (no-unauthorized-action) to
  every turn via a per-turn action log; day-loop grading (grade each email at the end
  of its serve day by attribution).
- **Cut (deliberately):** free-slot / "find an open slot" / `no_overlap` / can't-fit.
  Non-deterministic and the bulk of the complexity; conflicts are authored as
  fixed-time reschedules instead.
- **Maybe later (see `BACKLOG.md`):** per-seed ranged offsets (anti-memorization "random
  day"); global calendar + global grading; seed events (a pre-filled day-1 calendar).
- **Done:** webapp builder UI (time-in-the-builder, plain-language naming, obligation
  picker + auto-wired needle edges, Project Helios example) + re-vendor.
- **Pending:** grammar docs (document the `@HH:MM[-HH:MM]` time suffix in
  `ANSWER_KEY_GRAMMAR.md`).
