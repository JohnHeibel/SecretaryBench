# Saved-run audit continuation results

2/6 batches; 18/54 fresh checks; 16/48 emails; 32/96 initial reviewer assignments. New classifications resolved: 19 (including two reused pilot checks); unresolved new classifications: 1. 12 checks have queued interpretation, model-credit or saved-label decisions. 39/75 worksheet checks accounted for, including 19 preserved human judgments.

No benchmark scores were changed. AGENT_RESOLVED means coordinator-verified agent review, not human adjudication. Pending revisions to #14/#19 do not replace their saved human labels. Classification and model credit are separate; a KEY/PLAN finding does not establish model correctness.

## Coverage by batch

| Batch | Checks | Resolved | Needs James | Source spot-checks (seed 42) |
|---|---:|---:|---:|---|
| 1 | 8 | 7 | 1 | 22, B1 |
| 2 | 10 | 10 | 0 | 31, B4 |

## Classification totals

| Classification | Saved human | New agent-resolved |
|---|---:|---:|
| MODEL | 4 | 2 |
| KEY | 9 | 9 |
| PLAN | 4 | 1 |
| GRADER | 2 | 0 |
| AMBIGUOUS | 0 | 0 |
| PASS | 0 | 7 |

Sampled passes: 7 accepted for their targeted day-level check; 0 pending. Time/default discrepancies remain separately recorded. This selected-failure and sampled-pass audit is not a benchmark-wide accuracy estimate.

## Per-target findings

| Item | Status | Classification | Finding and model assessment | Evidence |
|---|---|---|---|---|
| 8 | NEEDS_JAMES | AMBIGUOUS | August 7 is explicit, but a mandatory calendar milestone is unstated. Both reviewers find the FYI and calendar interpretations defensible. No release event exists in the complete stores; only the separate keynote was added. | [review archive](batch-1-archive.json) |
| 21 | AGENT_RESOLVED | KEY | The rescheduling email does not revise the explicitly June 19 location-contact task or establish June 29 as its new deadline. No model pass for the entire email is implied. Leaving todo_22 unchanged is supported for this target; real-world contact completion is unknown. | [review archive](../agent-audit-pilot-sonnet-4-5.json#cases/C07) |
| 22 | AGENT_RESOLVED | KEY | The explicit athlete date is August 22; the key incorrectly requires June 22. evt_46 uses the correct day. Its 10:00-12:00 hours independently violate the 09:00, one-hour defaults. | [review archive](batch-1-archive.json) |
| 23 | AGENT_RESOLVED | KEY | The key requires a June 29 cutoff unsupported by the Friday request. The explicit turn-date rule supports July 3. todo_47 records selection and memo review by July 3 at 17:00. This supports the scheduling target, not actual review or communication. | [review archive](batch-1-archive.json) |
| 24 | AGENT_RESOLVED | PLAN | The June 29 delivery misses the explicitly pre-June-19 deadline; the exact June 19 key is also too restrictive. todo_49 substitutes unsupported July 3. Late-request handling and conflicting earlier Gold approval remain unresolved; no model pass is awarded. | [review archive](batch-1-archive.json) |
| 25 | AGENT_RESOLVED | MODEL | The explicit future July 15 move is supported and the existing board meeting is identifiable as evt_11. The request was read, but every event remains unchanged and evt_11 stays June 11. Historical hours are separate. | [review archive](batch-1-archive.json) |
| 26 | AGENT_RESOLVED | KEY | The key anchors preparation to July 22 despite the earlier explicit August 22 visit and conflicting next-week wording. todo_53 records the guest list due August 15. That fits August 22 but not next week; model credit remains unresolved. | [review archive](batch-2-archive.json) |
| 27 | AGENT_RESOLVED | KEY | Contacting invitees is a task before the visit, not an event exactly July 22. Both the keyed kind and date are unsupported. The combined todo_53 explicitly includes contacting invitees. Content is represented, but the conflicting visit date prevents a model pass; no actual outreach is established. | [review archive](batch-2-archive.json) |
| 28 | AGENT_RESOLVED | KEY | The exact July 15 deadline is unsupported by either the body’s pre-board window or a literal before-final reading of the subject. todo_54 is due July 21 before the July 22 board presentation, but after the July 15 final. Reviewers disagree on model credit; queued without overturning the definite key defect. | [review archive](batch-2-archive.json) |
| 29 | AGENT_RESOLVED | KEY | The requested pizza-party event move is to June 9; the key instead requires a July 14 task and selects unrelated todo_47. Delivery on July 5 is also stale. evt_7 remains June 8, with no replacement. Literal historical correction was not performed; unspecified stale-request handling prevents model-credit attribution. | [review archive](batch-2-archive.json) |
| 30 | AGENT_RESOLVED | MODEL | The observed July 7 request supports a July 14 sponsorship approval meeting. The key’s July 22 date is independently wrong. No corresponding meeting exists under any title in either full store; the model read the request with connected tools and omitted it. Failure remains justified after correcting the key. | [review archive](batch-2-archive.json) |
| 31 | AGENT_RESOLVED | KEY | The planning call explicitly names June 21, not the keyed July 21; July 8 delivery also makes it stale. No planning-call event exists or was created. Model credit depends on the unspecified treatment of historical requests, not on the invalid July key. | [review archive](batch-2-archive.json) |
| 32 | AGENT_RESOLVED | KEY | Forms are due before the June 21 call, not exactly July 21. The key has the wrong month and wrongly narrows a deadline window. No forms task exists or was created. A default 17:00 due time on the 09:00 call day would itself be late; delivery and missing backfill policy leave model credit unresolved. | [review archive](batch-2-archive.json) |
| B1 | AGENT_RESOLVED | PASS | The requested retreat date and event identity are satisfied by unique evt_4 on June 22. Day-level pass accepted. The eight-hour duration violates the one-hour default and is recorded separately. | [review archive](batch-1-archive.json) |
| B2 | AGENT_RESOLVED | PASS | Optional invitation supplies no date or scheduling request; no action is supported. Target was read; complete stores are identical, with the distinct July 5 budget meeting unchanged. | [review archive](batch-1-archive.json) |
| B3 | AGENT_RESOLVED | PASS | The HR update explicitly says nothing is needed. Target was read and no mutation occurred; full stores are identical. | [review archive](batch-1-archive.json) |
| B4 | AGENT_RESOLVED | PASS | The preview explicitly says Mark is not needed; no action is supported. Only unrelated keynote evt_17 is added; all previous objects remain unchanged, with no press-preview object under another title. | [review archive](batch-2-archive.json) |
| B5 | AGENT_RESOLVED | PASS | The explicit August 12 expo-keynote hold is uniquely represented by evt_17. Day-level pass accepted. The 10:00 start independently violates the 09:00 default; one-hour duration is correct. | [review archive](batch-2-archive.json) |
| B6 | AGENT_RESOLVED | PASS | The explicit August 6 launch dry run is uniquely represented by evt_20. Day-level pass accepted. Its 14:00-16:00 hours violate both the 09:00 start and one-hour defaults. | [review archive](batch-2-archive.json) |
| B9 | AGENT_RESOLVED | PASS | The date-only reschedule updates Serena evt_24 from September 11 to September 22 with no duplicate; Michael evt_25 remains unchanged. Targeted move satisfied. Earlier unsupported 10:00-12:00 hours are a historical observation, not a defect introduced by the move. | [review archive](../agent-audit-pilot-sonnet-4-5.json#cases/C08) |

See [NEEDS_JAMES](NEEDS_JAMES.md) for grouped decisions and [protocol](REVIEWER.md) for the closure rules. Original reviewer reports remain in `reviews/`; batch coordinator records contain source checks and spot-check findings.

Verification: all batch packet hashes, target coverage, Stage A-before-final ordering, source references, no future-email leakage, and original human text are checked by the private checkpoint validator. Existing Python suite: 107 passed. Webapp typecheck/lint and build are blocked by missing dependencies. No source, corpus, capture, grader, score, or production changes; no new benchmark runs.
