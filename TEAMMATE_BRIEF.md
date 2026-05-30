# Sprint 5: where we are and what's left (team brief)

Quick, plain-language summary so we're all on the same page before we finish Sprint 5. The detailed, file-by-file work order is in `SPRINT5_REMEDIATION.md`. You only need that one if you're implementing your part. This brief is the "why."

## The 30-second version

We migrated the AI lane from the Anthropic SDK to driving Claude Code (`claude -p`) as a subprocess. That part works. But the way it got wired, the benchmark's default run produces a score that's partly fake: calendar tasks can't pass, dates aren't actually checked, some criteria pass no matter what the model does, and we get no token logs. None of it is anyone's "fault" individually. It's a seam problem: the piece that runs by default is an incomplete copy of the piece that actually worked, and the grading upgrades that were supposed to go with it didn't land. The fix is straightforward and mostly mechanical.

## What's actually happening

There are two copies of "drive the model for one email":

- `model_runner.py`: the complete, correct one. Sets up a calendar, logs tokens, full instructions to the model.
- `harness.py` (the adapter): a thinner copy that the engine picks **by default**. It's missing the calendar setup, the logging, and half the instructions.

Because the engine always reaches for the adapter first, the good copy never runs. So on a normal `python engine.py` run:

1. The model is never told which calendar to use and can't create one, so **every calendar/meeting scenario scores 0.**
2. Criteria like "meeting on {date}" never get their date filled in before grading, so **the date is never actually checked** (right date and wrong date score the same).
3. Some criteria have no TC/CC/RS tag, and the grader **auto-passes** those. Worse, the criteria splitter is fragmenting a few real checks into auto-passing junk.
4. **No `token_usage.jsonl` / `tool_calls.jsonl`** are written, and compaction is never detected, even though measuring that was the whole point of the migration.

## Who this touches (not a blame list, just a map)

| Lane | Owner | Status | What's needed |
|---|---|---|---|
| Harness abstraction (`harness.py`) | Eyasu | merged, but the adapter is a partial port | restore calendar setup, full prompt, and logging (becomes a shared core) |
| AI-lane runner (`model_runner.py`) | Miguel | complete but bypassed | fold its working logic into the shared core; retire the duplicate |
| Engine + pooling (`engine.py`) | Nikita | works; one rough edge | clean up crash handling so a failed turn isn't silently counted as done |
| Grading (`grader.py`) | Anthony | the Sprint 5 grading items didn't land | resolve date tokens before grading, handle free-text criteria, fix the splitter, add token/compaction reporting |

## What we're proposing to do

1. **Collapse the two runners into one shared core.** This single change fixes most of the bugs at once and makes adding another model or harness easy later. One behavior, one place to fix things.
2. **Finish the grading upgrades** (Anthony's Sprint 5 items): actually check dates, decide how free-text criteria are handled, and report token/compaction info.
3. **Tidy the edges:** clean crash handling, make the run knobs consistent, do a docs truth-pass.
4. **Leave OpenRouter alone.** It's wired but unverified and off by default. It's intended scope for later, not part of getting the core working now. The shared-core change makes finishing it (and adding Codex) cheap whenever we want.

## A few facts worth knowing

- It currently **only works with `claude -p`.** The "any harness" design is a clean interface with one real implementation; Codex is a stub. Adding a second harness is a follow-on, not part of this.
- The test suite shows 6 failures + 1 error on a clean checkout, but those are **pre-existing and harmless**: 4 test a deleted email-delete feature, 2 need the server running, and 1 is a mis-named helper. 130 tests pass. We'll clean those up as part of this.
- The model swap (Haiku vs Sonnet vs Opus) already works. It's only the *harness* swap that's unbuilt.

## Next step

A fresh ultracode session will execute `SPRINT5_REMEDIATION.md`, which has every fix spelled out with file references, acceptance criteria, and an orchestration plan. There are four small decisions defaulted in that doc's "Execution Kickoff" section; if anyone has strong opinions on free-text grading or the scenario-id handling, now's the time to say so.

Questions or disagreements: bring them before we start Phase 0 (the shared-core refactor), since everything else builds on it.
