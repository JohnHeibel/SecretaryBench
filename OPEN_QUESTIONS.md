# Open Questions (parked)

Things we deliberately deferred while settling the grading model (`GRADING_MODEL.md`).
**Not decided yet** — revisit before/while building Phase 2 and the webapp.

## 1. To-dos: what's the best grading for deterministic outputs?

**How they're graded today** (`sb/grader.py`): a to-do `create`/`move` passes iff there
is **exactly one** to-do whose title+description contains the obligation's keywords,
**due on a day that satisfies the predicate**. `cancel` = none survive. A to-do is a
single point (its **due date**) — day-level, no duration like an event. Predicates:
`eq` (exact day) or `by` (on/before a deadline).

Things to decide later:
- **Exact day vs deadline as the norm.** `eq` for "do it that day," `by` for "by then."
  Confirm which is the default authors reach for, and whether both stay.
- **Should we grade the `completed` flag?** The store tracks it; the grader ignores it,
  so today we *can't* test "the model marked the prep task done." Adding it = carry
  `completed` into the grader's `Obj` + let an op assert it. Decide if that capability
  is worth the surface.
- **Day-level only, or can a to-do carry a time?** Leaning day-level ("due that day").
  Confirm.
- **Any determinism wrinkle vs events?** Don't think so (same machinery, simpler), but
  double-check when we revisit.

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
