# ADR 0001 — Replace `count` with verb-on-a-named-thing obligations

**Status:** accepted, staged · **Date:** 2026-06-01 · **Area:** `sb/grader.py`, `sb/schema.py`, corpus answer keys, webapp authoring surface

## Context

An answer-key `expect` entry today carries a `count` field. In the entire corpus it is
only ever `count: 1` (9 of 9 entries). Its job is to catch the classic reschedule failure:
the model creates the new event but forgets to delete the old one, leaving a double-booked
calendar. Without `count: 1`, a calendar with a stale duplicate still has one event on the
right day, so the entry would pass.

Two problems:

1. **`count` is hard to conceptualise.** Authors (and the maintainer) repeatedly asked "what
   does count do?" The number `1` is a stand-in for a boolean intent — *"there should be no
   duplicate of this thing"* — expressed as arithmetic. That mismatch is the confusion.
2. **It is the wrong primitive.** `count` is a patch over a missing concept: **object
   identity**. The grader matches calendar objects by fuzzy keyword (`title_match`), which
   has no notion of *which* event, so it counts them to detect duplicates. Once we name the
   thing an email acts on, "exactly one" is inherent and the count disappears.

This decision is also forced by a planned execution change (see
`DAY_LOOP_DESIGN_ISSUE.md`): the runner is moving from **one email per turn** to **one
simulated day per turn, model triages its own inbox**. When the model handles a batch of
emails in a single pass, calendar actions are no longer attributable to a single email, so
per-email / per-turn grading (and `count`, a per-turn-state patch) stops being well-defined.
The only thing that survives is grading the **resulting state of the world**.

## Decision

Move the answer-key model from *"assert how many objects match a keyword"* to
**"emit a verb on a named obligation"**, and grade by **state reconciliation**.

- **Obligation** — a named scheduling target (`acme_call`, `kickoff`) with a kind
  (event / todo) and a resolved date. It is the unit an author thinks in.
- **Verbs** — an email carries one of: `create` an obligation, `move` it (to a new date),
  `cancel` it. This is how a human narrates a calendar; "how many" never comes up.
- **Grading** — fold the verbs served so far into an expected live set, then reconcile
  against the model's calendar at the day checkpoint. Uniqueness, stale duplicates, and
  wrong-date all fall out of identity; no `count` annotation exists.

The DAG stays exactly as it is. `static` edges (depend on a non-date **fact** established
earlier) and `date` edges + anchors (depend on a **date** established earlier) are unchanged
and well-liked. Obligation verbs *reference* obligations and anchors, which means many edges
become derivable rather than hand-wired.

### Reconciliation scope (the one policy knob)

Per-obligation **closed**, globally **open**: for each known obligation's identity there must
be no unexpected/duplicate match (this recovers what `count: 1` did), but objects unrelated
to any obligation are ignored (so a benign prep block is not punished).

## What changes maps to what `count` did

| Old | New (reconciliation) |
|---|---|
| `count: 1` | one obligation = one object, inherently |
| `count: 0` / cancel | `cancel` the obligation; any surviving object fails |
| reschedule (create→move) | two verbs on the same obligation; final = one event on the new day |
| `forbid` (FYI email) | the email carries no verb; it adds nothing to the expected set |
| span (asked once, settled over N emails) | grade the final settled obligation, not each step |

## Staged migration

This is a wide change (grader, engine, oracle, analyze, live runner/store, the webapp builder
UI + seed, tests). It lands in stages so the suite stays green and the risk stays bounded.

- **Stage 1 — kill `count` from authoring (DONE, this ADR).**
  - `sb/grader.py`: cardinality now **defaults to exactly one**; `count` is read only to
    override (`0` = must not exist, `N` = exactly N). Behavior-preserving — every corpus
    entry was already `count: 1`.
  - Stripped the 9 redundant `count: 1` lines from `corpus/nodes/*.json`.
  - Updated `ANSWER_KEY_GRAMMAR.md` and `webapp/AUTHORING_GUIDE.md` to document the default.
  - The `count` field stays in the schema for the `0` / `N` exceptions (e.g. `oracle.py`
    uses `count == 0` to emit a cancellation).
  - `pytest sb/tests/` stays at 45 passing.
- **Stage 2 — verb authoring surface.** Add `op: { create | move | cancel: <name>, ... }` to
  the schema as the authored form; compile it down to expected-state for the existing grader.
  Update the webapp builder so authors pick a verb + a named thing instead of an `expect`
  block. Land alongside or just before the day-loop change.
- **Stage 3 — reconciliation grader + day loop.** Replace per-email/per-turn grading with
  checkpoint state reconciliation (`DAY_LOOP_DESIGN_ISSUE.md`). At this point the per-email
  `expect`/`forbid` path and the residual `count` field can be retired.

## Consequences

- **Authors stop touching `count`.** The grammar's vocabulary becomes *create / move / cancel
  a named thing* plus dates — far easier to teach to club authors, which is the whole point
  of the webapp.
- **No-double-book is the default**, so latent author mistakes (forgetting `count: 1`) can't
  silently weaken a test.
- **`count: 0` and `count: N` remain expressible** during the staged transition; cancellation
  gets a first-class verb in Stage 2.
- The destination is coupled to the day-loop execution change; Stages 2–3 should not land
  before that work is ready.

## Alternatives considered

- **Keep `count`, apply it consistently everywhere.** More explicit, but doubles down on a
  primitive authors find confusing and that the day-loop change invalidates anyway.
- **Rename `count: 1` to a `unique: true` boolean.** Fixes the naming confusion cheaply, but
  still bolts cardinality onto a keyword match instead of giving the thing an identity. Good
  enough if we were *not* rebuilding; we are.
