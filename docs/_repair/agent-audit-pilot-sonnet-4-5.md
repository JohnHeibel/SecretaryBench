# Agent-assisted audit pilot: saved Sonnet 4.5 run

Date: 2026-09-25. Status: **pilot complete; proposed findings await James's adjudication**.

This is a one-time assisted manual audit of the saved run. It does not implement or
replace the benchmark grader. No corpus, grader, capture, score, or original worksheet
verdict was changed.

## Result

Six previously reviewed checks and two unreviewed checks received two fresh agent
reviews each. Both reviewers chose the same primary category on **8/8 cases**. They
matched James's saved category on **4/6 previously reviewed cases** and challenged two.
These are agreement counts, not accuracy estimates. The sample was selected for
coverage, and the reviewers used the same inherited model configuration and rubric.

| Worksheet item | Target | Saved James verdict | Reviewer 1 | Reviewer 2 | Evidence-based finding |
|---|---|---|---|---|---|
| 6 | Outsole vendor follow-up | KEY | KEY | KEY | June 12 is the delivery deadline. A June 11 contact reminder preserves it; requiring contact exactly June 12 is too restrictive. |
| 9 | CTO meeting | GRADER | GRADER | GRADER | The model created the correct June 19 meeting as `Meeting with CTO`; the hidden obligation-name matching rejected it. |
| 14 | Choose a podcast slot | KEY | MODEL | MODEL | The email explicitly says `pick one` of three offered slots. The model recorded a reminder to decide but booked none. |
| 16 | Approve prize amounts | KEY | KEY | KEY | Nothing specifies the key's June 25 deadline. The model's June 24 reminder is also discretionary, so full model compliance remains ambiguous. |
| 19 | Site launch milestone | MODEL | KEY | KEY | The email fixes the launch date without explicitly requesting a calendar hold. Instructions permit no action on FYIs/confirmations. Mandatory event creation is disputed. |
| 20 | Vision sync | PLAN | PLAN | PLAN | A June 15 meeting request arrived June 27. Whether to backfill an expired request is unspecified. |
| 21 | Move retreat-location contact task | Unreviewed | KEY | KEY | The email moves the retreat, but does not revise the contact task's June 19 deadline or establish a June 29 replacement. |
| B9 | Reschedule Serena visit | Unreviewed | PASS | PASS | The model updated the existing Serena event from September 11 to September 22 without a duplicate or changes to Michael's event. |

KEY means an unsupported or wrong key; GRADER means a rejection of correct work
despite a supported requirement; MODEL means a supported requirement was missed;
PLAN means a delivery-timing defect. PASS here applies to the targeted check.

The two new judgments are proposals. The original worksheet still has **19 checked
verdicts**, through item 20 with item 8 blank. It still has 41 unchecked failed checks
and 15 unchecked sampled passes. This pilot has not silently promoted agent findings
to human-approved labels.

## Judgments to revisit

### Item 14: who chooses the podcast slot?

Both agents interpreted `Any of these afternoons works, pick one` as discretion to
book August 3, 4, or 5 at 15:00. Their pre-key assessments already reached that reading.
The complete calendar shows no conflict in those slots. The trace shows only a
`Pick podcast taping date` task due June 21, with all three options still unselected
([day 19 trace](../../captures/baseline-sonnet-4-5/raw/day_019.jsonl)).

**Coordinator assessment:** MODEL is better supported by the supplied instructions.
James's saved KEY judgment assumed the principal should choose. That authority
restriction is not stated to the evaluated assistant. This is a proposed correction,
not an automatic overwrite of the saved label.

### Item 19: does a fixed launch announcement require a calendar event?

Both agents classified the missing milestone as KEY with **medium confidence**. The
email says the site launch is official and its date fixed; it separately asks for
sign-off by August 3. The model created the sign-off reminder. The instructions allow
no action for FYIs/confirmations and say to avoid over-acting. No rule says to record
every dated business milestone ([day 26 trace](../../captures/baseline-sonnet-4-5/raw/day_026.jsonl)).

**Coordinator assessment:** the current evidence does not justify a confident MODEL
failure. Recording a launch milestone is useful and reasonable, but requiring it
needs a clearer task contract. James should adjudicate KEY versus AMBIGUOUS before
changing the existing label. Agreement between the agents does not resolve the
intended policy by itself.

### Other interpretation boundaries

- **Item 16:** an unsupported exact key date does not prove the model's alternative
  date satisfies an undated announcement deadline. Both reviewers distinguished
  key invalidity from model correctness.
- **Item 20:** the requested meeting is stale, but writing a historical calendar
  entry is technically possible. The evidence supports a PLAN defect, not a claim
  that the tools made every possible response impossible. Backfilling policy is
  unspecified.
- **Item 21:** a new venue check could be sensible after rescheduling. The available
  instructions still do not establish the key's exact June 29 task deadline.
- **B9:** the correct date-only move inherited the earlier event's nondefault hours.
  That separate original-creation issue does not overturn this rescheduling check.

## What the pilot establishes

The agents produced traceable, bounded reviews, confirmed known failure modes, and
challenged prior interpretations in both directions. This supports using agents to
draft the remaining audit, with human adjudication of disputed requirements and
spot-checks of agreements. It does not establish human-equivalent accuracy or a
population error rate.

Before scaling, keep **target findings**, **sibling observations**, and **earlier
model errors** in separate fields. Some reviewers placed unrelated observations in
`secondary_categories`. The raw reports preserve those observations, but they are
excluded from the agreement counts and target classifications above. Also preserve
the distinction between an invalid key and a demonstrated correct model response.

No additional audit batch is authorized by completion of this pilot. The next step
is James's adjudication of the eight results, especially items 14 and 19, followed by
a decision about the remaining audit's scope and review coverage.

## Method and evidence

- Baseline: `captures/baseline-sonnet-4-5`, certified model
  `claude-sonnet-4-5-20250929`, corpus hash `03e0d963b9866d8f`.
- Source checkout: `debug/model-resolution` at `540c76ec10c68e36b9b2e11a1bd61ab77aa43a5b`.
  The worksheet uses the current re-grader's 114/167 result on that saved capture;
  the original as-run result was 97/167. Neither score changed.
- Selection: items 6, 9, 14, 16, 19, 20, 21 and B9. Seven currently failed checks and
  one sampled pass. Deliberate coverage, not random sampling.
- Sixteen fresh reviewer contexts, no forked conversation, maximum three concurrent.
  No reviewer saw James's saved labels or the other reviewer's output. This was
  procedural isolation, not a filesystem access boundary. The exact reviewer-model
  identifier was not returned by the collaboration tool; no model overrides were used.
- Each reviewer saved an initial obligation/date interpretation before reading the
  target key, grader verdict, model actions, or end-of-day state. All 16 reported no
  later change to that initial interpretation. This self-report is not a proof that
  later evidence had no influence.
- Inputs included the benchmark's custom instructions, rendered target and earlier
  emails, same-day context, complete unfiltered before/after stores, and raw tool
  requests/results. Prior inbox and mutation history were available. Future emails,
  model thinking blocks, and human verdicts were excluded.
- Custom instructions match pre-capture runner commit `d396921`. The complete serve
  plan and corpus hash match the capture. Every target body also matches its actual
  captured `get_email` response.

[The review archive](agent-audit-pilot-sonnet-4-5.json) contains the common protocol,
source provenance, supplied Stage A/Stage B context, all 16 initial interpretations,
all 16 final reports, packet checksums, and mappings to the existing capture files.
Large state and trace data remain referenced in the committed capture. Complete
original packets and working notes are also retained locally in
`.Codex/plans/audit-pilot-2026-09-25/`.

James's saved labels are human-approved and sometimes AI-assisted. They are useful
comparison references, not independent unaided human gold. The coordinator had seen
those labels and is not a seventeenth blind reviewer.

## Verification

- All 56 packet files retain their initial checksums; all 16 final reports parse,
  identify the expected case/reviewer, and cite evidence.
- All 16 initial-interpretation files predate their final-report files. The original
  worksheet's SHA-256 is unchanged.
- Existing Python suite: `PYTHONDONTWRITEBYTECODE=1 .venv/bin/python -m pytest sb/tests -q -p no:cacheprovider`
  returned **107 passed in 3.29s**. This checks existing software behavior, not the
  correctness of the audit judgments.
- Typecheck/lint/build were not run because `webapp/node_modules` is absent. No
  application or UI code changed, and dependencies were not installed for this audit.
- No live benchmark run, production access, corpus edit, grader edit, or verified
  repair-status promotion occurred.
