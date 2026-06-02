# ADR 0001 — Replace `count` with verb-on-a-named-thing obligations

**Status:** accepted · stages 1–2 landed · **Date:** 2026-06-01 · **Area:** `sb/grader.py`, `sb/schema.py`, corpus answer keys, webapp authoring surface

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

This decision is also aligned with the day-loop execution change (historically tracked in
`DAY_LOOP_DESIGN_ISSUE.md`): the live runner moved from **one email per turn** to **one
simulated day per turn, model triages its own inbox**. When the model handles a batch of
emails in a single pass, the durable grading primitive is the **resulting state of the
world**, not the exact edit sequence the model used.

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

### Obligation identity = a node-scoped anchor (decided: "Option A")

The obligation name *is* an anchor. `create: kickoff` registers obligation `kickoff` and, when
a sibling op references `@kickoff`, publishes `@kickoff` = its resolved date. So `move: kickoff`
references it by name, its dependency edge to the creator **derives**, and its date is reusable
as `@kickoff` (no repeating the base expression). This is strictly *less* to author than the old
model (which repeated title keywords, edges, and base dates) — the win that motivated the change.

Names are **node-scoped**: the loader qualifies them internally (`__obl_<node>__<name>`), so
`globex-acq` and `henderson` can both own a `kickoff` without colliding in the global anchor
table. Authors only ever write the bare name. Cross-node references still use body anchors
(`{!signing=…}`), the pre-existing mechanism.

### Reconciliation scope (the one policy knob)

Per-obligation **closed**, globally **open**: for each known obligation's identity there must
be no unexpected/duplicate match (this recovers what `count: 1` did), but objects unrelated
to any obligation are ignored (so a benign prep block is not punished). No-action emails
remain strict: `ops: []` means the model must create nothing attributable to that email, so
bait/FYI emails stay discriminating.

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
- **Stage 2 — verb model + reconciliation grader (DONE).**
  - `sb/schema.py`: the answer key is now `ops: [ { create|move|cancel: <name>, kind, on,
    match } ]`. Obligations are node-scoped anchors (Option A above): `create` publishes
    `@name` when referenced; `move`/`cancel` inherit the obligation's kind/match and
    auto-derive their dependency edge to the creator. `expect`/`forbid`/`count` are gone.
  - `sb/grader.py`: per-obligation reconciliation against cumulative node state —
    `create`/`move` = exactly one match on the `on` day, `cancel` = none. No `count`.
  - `sb/oracle.py`, the e2e + schema tests, and all five corpus nodes migrated (hard cutover,
    no dual-read path). Oracle is 100% on the migrated corpus; `pytest sb/tests/` at 46.
  - Webapp builder migration to the verb form: in progress (vendored `sb` + the React
    authoring surface).
- **Stage 3 — day loop + live runner integration (DONE).** Serving moved from one-email-per-turn
  to one-day-per-turn in `sb/live/`: the model lists the day's inbox, reads each email, acts
  through MCP tools, and the runner grades the resulting state at the day checkpoint. Created
  objects still carry `email_id`, which preserves strict no-action grading for bait/FYI mail.

## Consequences

- **Authors stop touching `count`.** The grammar's vocabulary becomes *create / move / cancel
  a named thing* plus dates — far easier to teach to club authors, which is the whole point
  of the webapp.
- **No-double-book is the default**, so latent author mistakes (forgetting `count: 1`) can't
  silently weaken a test.
- **`count: 0` and `count: N` remain expressible** during the staged transition; cancellation
  gets a first-class verb in Stage 2.
- The destination is coupled to the day-loop execution change now implemented in `sb/live/`.

## Alternatives considered

- **Keep `count`, apply it consistently everywhere.** More explicit, but doubles down on a
  primitive authors find confusing and that the day-loop change invalidates anyway.
- **Rename `count: 1` to a `unique: true` boolean.** Fixes the naming confusion cheaply, but
  still bolts cardinality onto a keyword match instead of giving the thing an identity. Good
  enough if we were *not* rebuilding; we are.
