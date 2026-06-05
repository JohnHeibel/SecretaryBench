# Open Questions (parked)

Things we deliberately deferred while settling the grading model (`GRADING_MODEL.md`).
**Not decided yet** — revisit before/while building Phase 2 and the webapp.

## 1. To-dos: what's the best grading for deterministic outputs?

**How they're graded today** (`sb/grader.py`): a to-do `create`/`move` passes iff there
is **exactly one** to-do whose title+description contains the obligation's keywords,
**due on a day that satisfies the predicate**. `cancel` = none survive. A to-do is a
single point (its **due date**) — day-level, no duration like an event. Predicates:
`eq` (exact day) or `by` (on/before a deadline).

**Resolved 2026-06-04 — the settled design:**
- **Grade on the due date, day-level.** A to-do is a single point (its due date); the model's
  free-text title/description is *never* judged for meaning, only matched for the keyword
  (attribution). So the grade is a fixed function `(exactly one to-do whose title contains the
  keywords) AND (due date satisfies the predicate)` — deterministic, same machinery as events
  minus duration. **No determinism wrinkle.**
- **`by` is the default** ("do it by Friday" — any day up to the deadline passes, a fixed set),
  with `eq` for "that exact day." Both stay.
- **No clock on to-dos.** The builder hides the time control for to-dos; a due *time* adds no
  deterministic signal and only complicates attribution. (Day-level confirmed.)

Still open:
- **Grade the `completed` flag?** Deferred. The store tracks it and the grader ignores it, so we
  can't yet test "the model marked the task done." Worth adding only paired with a deadline
  ("complete by Tue") so it still exercises timing; otherwise it's model-action surface for little
  temporal gain.
- **Keyword attribution at real-author scale** — the live worry. See §2.

## 2. Obligation keywords: make them painless for authors

**What they are:** the words the grader uses to tie the model's calendar object to an
obligation (substring match on the title the model chose). **The model never sees
them.** They default to the obligation's name.

**The problem:** "match keywords" + the collision lint is confusing for club authors.
Authoring should feel like writing an email, not programming a matcher.

Ideas to explore (mostly webapp, but #5 is a core idea):
1. Present as **one plain field**: "What's this event called?" — never "match keywords."
2. **Auto-fill from the name / email subject**; the author usually leaves it alone.
3. **Hide keyword tuning under "Advanced"** (95% of authors never open it).
4. **Auto-suggest** keywords from the email body/subject.
5. **Reduce the keyword dependency at the source.** Every object already carries the
   `email_id` that created it, and obligations are node-scoped — so could the grader tie
   objects to obligations primarily by **attribution + kind**, with keywords only a
   fallback to disambiguate *multiple objects from the same email*? Worth a hard look:
   it could remove keywords from the common case entirely.
6. Cross-scenario keyword collision = a soft **warning**, never an error (the thread
   fence already guarantees correctness; distinct naming is nice-to-have, not required).

When we revisit: pick the smallest change that makes naming invisible to authors in the
common case while keeping grading unambiguous.

**Scale risk (2026-06-04, the live worry).** Keyword attribution is validated only against the
oracle (which titles objects *exactly* with the keywords) and a couple of hand-authored scenarios.
With ~20 real authors writing real emails — and real models choosing their own titles — keyword
alignment is **untested at scale**. Today it holds via the oracle satisfiability check + the
collision lint + forgiving substring matching, but that is "works in practice," not a guarantee.
**The hardening is idea #5: attribute primarily by `email_id` + `kind`** (the tag the harness sets
on every object, not the model's prose), with keywords only a fallback to split multiple same-kind
objects from one email. That removes the model's word-choice from the common case entirely. It's a
small grader change (the `Obj` already carries `email_id`); do it before the corpus gets large.
