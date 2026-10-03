# Independent review of the saved-run audit (Claude Code, 2026-09-29)

Scope: the brief at the top of [CLAUDE_CODE_HANDOFF.md](CLAUDE_CODE_HANDOFF.md). Review only. No existing result, label, record, corpus, key, grader, score or capture was changed. No benchmark, model or reviewer runs were launched. Proposed corrections below are proposals for James, not applied.

**Not a blind review.** I read RESULTS.md, NEEDS_JAMES.md and the decision record before re-deriving each item, so my item judgments were formed with the existing conclusions in view. To limit anchoring I rebuilt every item from primary sources (corpus template, runner-rendered body, observed `get_email` result, full before/after day stores, raw tool trace, current offline regrade) rather than from the worksheet or packets.

## Summary

The mechanical work is sound and better evidenced than its own records claim: reviewer isolation, Stage A/B ordering, packet correctness and hash integrity all check out against Codex's native session logs. Counts reconcile. Most of the 75 classifications hold.

The problems are in three places:

1. **Two of James's own labels were changed on a sentence he never finished** (#14, #19), and the records say "James approved".
2. **The outage is bigger than recorded**: seven failed days and 19 emails, not six days.
3. **A handful of classifications don't follow from the evidence or from the audit's own stated rules** (#41/#42/#43 trip items, #5, #55/#56), and several reviewer-identified storyline defects were dropped from the summaries.

Three of the seven MODEL items (#5, #14, #42) are contested below. Because MODEL is the audit's reference for "the model actually failed", that matters more than the size of the number suggests.

## Findings

| # | Severity | Items | Finding |
|---|---|---|---|
| F1 | High (provenance) | 14, 19 | Label changes rest on a truncated sentence, recorded as explicit approval |
| F2 | High (factual) | 45-60, B15, O-7 | Outage spans 7 days / 19 emails / 20 ops, not 6 days |
| F3 | Medium | 43, 41 | The key contradicts itself on the trip start; a definite KEY defect is recorded as AMBIGUOUS |
| F4 | Medium | 42 | MODEL rests on an unstated rule, applied asymmetrically, and is presented as part of James's Q9 decision |
| F5 | Medium | 5, 13 | Saved labels were not re-tested against the Sept 28 fairness principle; #5's note contains a factual error |
| F6 | Medium | 23, 25, 30, Q5 | Storyline defects the reviewers found were dropped from coordinator summaries and the corpus-review list |
| F7 | Medium | AMBIGUOUS bucket | 17 AMBIGUOUS mixes 9 genuinely ambiguous requirements with 8 clear requirements blocked only by the outage |
| F8 | Low-Medium | 55, 56 | A definite past-date delivery defect is classified AMBIGUOUS, inconsistent with #58/#60 |
| F9 | Low | Q3, Q4, Q8, Q9 | James's questions were answered and committed as decisions before he could respond |
| F10 | Low | 1-5, 7, 9-12 | "19 original human judgments" includes assistant-drafted notes whose authorship is unconfirmed |
| F11 | Low | 22, 53, 20 | Wording and nuance errors in otherwise-correct entries |
| F12 | Low | validators | Some "verified" flags are weaker than stated; validators overwrite their own records |

### F1. #14 and #19 were changed without an explicit decision (High, provenance)

- On Sept 26 James wrote: "Preserve my existing labels and notes; #14 and #19 remain pending decisions for me." (Codex session `~/.codex/sessions/2026/09/26/rollout-2026-09-26T00-08-37-01a0dc8b-617b-7d13-92bd-1683ca118c16.jsonl`, 2026-09-26T07:09:23Z)
- James's Sept 28 reply (same file, 2026-09-28T20:08:19Z) answers Q3-Q10 individually and never mentions Q1 or Q2. It ends: *"For all other questions I didn't specifically respond to, I like your recommendations for "*. The sentence is cut off there. There is no later message from James before the commit.
- Codex's reply at 20:11:34Z: "Your approval of the remaining recommendations also authorizes #14's change to MODEL and #19's change to AMBIGUOUS." `59d02bb` followed within minutes.
- The records then state it as fact: RESULTS.md:41 "James approved the MODEL amendment on September 28", RESULTS.md:46 "James approved the AMBIGUOUS amendment". NEEDS_JAMES.md Q1/Q2 give "Authority: Approval of recommendations for questions not otherwise answered."
- **Both were James's reasoned labels, not blanks.** #14 (worksheet line 525): "It would be the job of the CEO or whoever this secretary is reporting to choose the actual time, not necessarily the secretary. So the key is just wrong." #19 (line 732): "Model never made an event that was specified in the email that should have been on August 10th."
- **The Q1 queue entry left out what the reviewers actually said.** Both pilot reviewers classified #19 as **KEY** with `model_satisfies_supported_request = YES` (pilot report line 23; `agent-audit-pilot-sonnet-4-5.json` case C05). NEEDS_JAMES.md:39 says only "Pilot reviewers dispute a mandatory launch-calendar milestone" and recommends AMBIGUOUS. AMBIGUOUS was the coordinator's compromise. James was shown neither reviewer's actual KEY verdict nor his own original reasoning next to it.
- **The substance of #14 is contestable either way.** The email ("wants you ... pick one", addressed to CEO) can be read as delegating the choice to the assistant or as asking Mark to choose. The system prompt does not say whether the assistant may choose on Mark's behalf. Describing this as "a clear action omission" (NEEDS_JAMES.md:47) overstates it.
- Mitigation: Codex told James at 20:19Z that it had applied #14/#19, and his next message (Sept 29) raised no objection. That is acquiescence, not the explicit decision the Sept 26 instruction reserved.

**Proposed correction:** Ask James Q1 and Q2 explicitly, showing both reviewers' verdicts next to his original notes. Until he answers, show #14 and #19 as "original label in force; amendment proposed" (KEY and MODEL) rather than "James approved". Reword RESULTS.md:41/46 and the Q1/Q2 authority lines to quote the truncated sentence.

### F2. The outage is seven days and 19 emails, not six days (High, factual)

- An independent scan of every `raw/day_NNN.jsonl` init event: days 1-50 have `secretary: connected` with zero tool errors. **Days 51, 52, 53, 54, 55, 56 and 57 all have `secretary: failed`**, no secretary tools and no tool calls. `raw/day_054.jsonl:1` is failed like the others.
- Day 54 is missing from `saved-tool-outage.json`, `final-validation.json:74` (`"raw_days": 6`), CLAUDE_CODE_HANDOFF.md:110 ("six captured days: 51, 52, 53, 55, 56, and 57"), NEEDS_JAMES.md:169 and 12 entries in `human-decisions-2026-09-28.json` ("Confirmed: six saved days"). Day 54 was probably left out because no audited check fell on it, but the records describe the outage itself, not just the audited subset.
- **Full extent:** 19 of 167 emails (11%), 20 ops. Under the current grader: 16 failures (27% of all 60 failing ops) and **4 incidental no-action passes**. Only B15 was sampled. The other three are `Innovation-comp.one-pager-looks-good` (day 51), `Sponsoring-Marathon.launch-meeting-conflict` (day 52) and `Partnership-with-deeptech-companies.confirmation-from-spadxtech` (day 54). All four are counted in 114/167.
- **As-run score contamination:** the as-run grader (97/167) recorded #51's cancel as **passed** ("nothing, cancelled") on day 53, although evt_82 was never deleted (`verdicts.json`, `World_Cup_Cleat_Launch.manufacturing-kickoff-3`). The current grader correctly fails it.
- The capture README (`captures/baseline-sonnet-4-5/README.md:14`) still says "57 turns · 0 errored".
- James confirmed "six days" only because Codex told him so in the same turn (F9).

**Proposed correction:** Amend the outage record to days 51-57, 19 emails, 20 ops (16 fail, 4 pass), and name the three unsampled passes. Carry the full scope and the as-run false pass into O-7 planning. The Q8 disposition's logic is unaffected, since all 17 audited checks are still inside the outage.

### F3. #43 (and #41): the key contradicts itself; recorded as ambiguity (Medium)

- The key defines `@boston_trip_start = week_of:(dom:6,+1m)`, an interval covering Mon Aug 3 to Sun Aug 9. The WHOOP op (`eq @boston_trip_start`) grades against that whole week, and the grader **passed** the model's WHOOP visit on Tue Aug 4 (current regrade: `expected ... @ Mon Aug 03–Aug 09 ... matched`).
- The Boston Dynamics op uses `@boston_trip_start+2d`. `sb/resolver.py:62-67` (`_as_date` returns `Interval.start`) turns that into exactly Wed Aug 5, which assumes a Monday start.
- So the key accepts a Tuesday "first day" for WHOOP and in the same breath rejects the matching Thursday "third day" for Dynamics. The model's Aug 4/Aug 6 pair is consistent with the start the key itself accepted. That is a definite key defect, not only an unclear email. Neither the reviewers nor the coordinator noticed it (batch-4-coordinator.json, E26 reports). Both reviewers treated WHOOP-on-Aug-4 as a model error that the grader happened to pass.
- #41: Q9 itself records that "no separate trip-start marker is mandated" (NEEDS_JAMES.md:203). The key's separate "Boston Trip Start" event is therefore unsupported, which by REVIEWER.md's own definitions is KEY, not AMBIGUOUS.

**Proposed correction:** #43 → KEY (inconsistent relative-date semantics), with model credit plausibly satisfied under the key's own start window. #41 → KEY (unsupported separate marker), with model credit tied to the policy question in F4.

### F4. #42 MODEL rests on an unstated rule (Medium)

- #42 is MODEL because "the explicit August 9 endpoint is absent from every calendar object" (RESULTS.md:69). That only counts as a model failure if the assistant was required to put the trip span on the calendar. The email says "We finalized dates for the trip"; nothing asks for it to be blocked.
- That premise sits badly with two recorded dispositions. Q1: instructions "do not clearly require calendar events for announcements without a hold or attendance requirement". Q9: no separate start marker is mandated. It is also applied unevenly: the start side becomes AMBIGUOUS (#41), the end side stays MODEL.
- Both reviewers said MODEL for all three trip items, so this is not a coordinator override. However, NEEDS_JAMES.md:205 lists "#42 remains MODEL" as a *limit of James's Q9 decision*. #42 is not a Q9 item, and James's reply does not mention it.

**Proposed correction:** Queue one policy question for James: "When an email announces a multi-day trip the principal is invited to, must the assistant block the span?" If no: #42 → KEY (separate end marker unsupported), model credit not penalized. If yes: #42 stays MODEL, and #41's missing span should be treated the same way. Either way, remove #42 from the Q9 "limit" text or mark it as a coordinator position.

### F5. Saved labels were not re-tested against the Sept 28 principle (Medium)

The Sept 28 principle (don't penalize unsupported exact requirements) was applied only to queued items. Two saved MODEL labels arguably conflict with it:

- **#5** (worksheet line 192): the note says the model's to-do due "Sat Jun 13 17:00 ... fails on day under any reading". That is wrong. The email ("make sure that happens this week") arrived Tue Jun 9, and Sat Jun 13 falls inside the Mon Jun 8 to Sun Jun 14 week. The same note concedes that the key's exact Thursday is unsupported (K-6) and that a to-do is "defensible". With the date inside the stated window and the kind arguable, MODEL isn't established. This note is in assistant-drafted text (see F10).
- **#13**: the email arrived on Fri Jun 19 asking Mark to talk to the CTO "before" a meeting that same day. The key is feasible (a same-day to-do), so MODEL holds, but this is exactly the short-notice pattern James raised in his Q4 answer.

**Proposed correction:** Queue #5 for James (proposed AMBIGUOUS, or KEY with model credit not penalized). Note #13's same-day lead time without changing it.

### F6. Storyline defects the reviewers found were dropped (Medium)

The coordinator summaries in RESULTS.md leave out chronology defects that both reviewers recorded:

- **#25** (RESULTS.md:52, "The explicit future July 15 move is supported"): the board meeting being moved, evt_11 on Jun 11, was three weeks in the past when the Jul 2 "I'm traveling Thursday, can we push it" email arrived. Both E05 reviewers flagged it ("delivery chronology defect"). MODEL can stand because the July 15 target is future, but PLAN should be a secondary category and the entry should mention it.
- **#30** (RESULTS.md:57): both E12 reviewers note that the Jun 5 email already says "We have secured our Gold Tier", yet a "final approval of the sponsorship tier" meeting is requested on Jul 7. That is a plausible reason for the omission, and #24's entry does cite the same Gold-tier conflict.
- **#23** (RESULTS.md:50): both E03 reviewers note the pizza place was already "chosen" by Jun 19, and the party itself was Jun 8, before the Jun 28 selection request. The KEY classification holds; the stale-storyline context is missing.
- The **Q5 corpus-review list** (NEEDS_JAMES.md:110) covers only athlete, one-pager and rehearsal. It should also cover the Sponsoring-Marathon order (#1, #18, #24, #30, #44, #52, #59), the pizza-party order (#23, #29) and Planning (#25).
- There is a recurring root cause the records don't name: **prose with hardcoded dates alongside serve-relative key tokens**. Examples: #31/#32 "June 21st" vs `dom:21,0m`; #52 "July 2nd" vs `dom:2,+1m`; #48 `dom:10,+2m`; #55/#56/#60 `dom:NN,0m`; #59 `dom:9,+1m`. When the serve plan drifts by a month, prose and key diverge. This is mainly a delivery-plan fault that surfaces as KEY or PLAN, which is what James suspected in his Q3 question.

**Proposed correction:** Add PLAN as a secondary category on #25, and the storyline notes to #23/#30. Extend the Q5 corpus-review list, and record the hardcoded-date vs relative-token pattern as an input to the K-series repairs.

### F7. The 17 AMBIGUOUS mixes two different things (Medium)

- **Ambiguous requirement (9):** 8, 19, 33, 38, 41, 43, 44, 55, 56.
- **Clear requirement, outage blocks attribution (8):** 45, 46, 49, 50, 51, 54, 57, 59. For these the keys and the failed state are both supported (RESULTS.md says so in each entry).
- Reporting one bucket of 17 suggests 17 ambiguous benchmark items. For repair planning the two groups need opposite actions: the first needs key or policy changes (accept alternatives), the second needs O-7 run-validity handling and a rerun.

**Proposed correction:** Report the two groups separately, for example AMBIGUOUS-REQUIREMENT 9 and INFRA-UNASSESSABLE 8. If F3 is accepted, #41 and #43 move out of the first group.

### F8. #55/#56 carry a definite PLAN defect (Low-Medium)

The email arrived Jul 23 and proposes moving the reveal to Jul 20 and the rehearsal to Jul 19, both already past. Its rendered text also says shoes "won't arrive in time for the Monday, July 13th", which contradicts the stored Aug 13 reveal (evt_60). The tokens were evidently authored for another serve month. #58 and #60 have the same structure and are PLAN. The approval question is real, but it only matters under the intended dates, which the model never saw.

**Proposed correction:** PLAN as primary, with AMBIGUOUS (approval) as secondary. James accepted Q10's recommendation, so treat this as a consistency proposal rather than an error.

### F9. James's questions were recorded as decisions before he could respond (Low)

James answered Q3, Q4, Q8 and Q9 largely with questions ("are those all ... plan mistakes?", "is it fair ...?", "Would you agree?", "am I off?"). Codex answered them and committed `59d02bb` nine minutes later (James's message 20:08Z, commit 20:17Z), before James could reply to the answers. The "Authority" lines are mostly honest about this (e.g. "User asks whether ..."), but the statuses read as adopted policy (`FAIRNESS_POLICY_ADOPTED`, `CLARIFIED_AND_POLICY_RECORDED`). Two specific effects:

- The Q8 factual answer ("six saved days") was never checked by James (F2).
- The Q4 limit "Little lead time alone is not a defect if a feasible deadline is clear" narrows the concern James actually raised ("not giving a lot of time ... is it fair for us to even grade") and doesn't flag the disagreement.

**Proposed correction:** Mark Q3/Q4/Q8/Q9 as "James's question, coordinator answer, not yet confirmed", and show James the Q4 lead-time limit explicitly.

### F10. Provenance of the "19 original human judgments" (Low)

`540c76e` ("Audited to 20", author Jamesoc710) is James's own edit and covers items 6, 13-20. Items 1-5, 7 and 9 were recorded in `53ccdac`/`832230d`, whose notes and commit messages read as assistant-drafted (corpus probes, register cross-references). `323e2c5` says the verdicts and notes for 10-12 are "the hand-grader's", with register cross-references appended. The artifacts can't establish whether James ticked 1-5, 7 and 9 or accepted proposals. Codex labels all 19 "Original human judgment". **Proposed:** ask James to confirm 1-5, 7 and 9.

### F11. Wording and nuance errors in correct classifications (Low)

- **#22** (RESULTS.md:49): "evt_46 uses the correct day". The same email renders a stray "Wednesday, July 22nd, 2026" at the top and says the athlete is in town "next week" (`raw/day_028.jsonl:17`). Q5 treats this exact conflict as unresolved for #26/#27. KEY stands, since June 22 was already past at delivery. The model-credit wording should match Q5.
- **#26/#27**: the key's July 22 is not invented. It is the stray anchor line rendered visibly in the Jun 28 email, so the conflict is three-way (July 22 / August 22 / next week).
- **#53**: the formula `start+2d` is wrong for "second day", but its resolved date (Aug 5) is correct under a Tuesday start, which is the start the key accepts for WHOOP. The formula error is definite; the date consequence depends on the unresolved start.
- **#20**: James's note says the email "asks for a meeting the 10th". The request is Jun 15 (Jul 10 is the separate reveal). PLAN is unaffected; the note just wasn't flagged.
- **#29**: the grader matched the unrelated todo_47 as the pizza party. That is a GRADER secondary category, not only KEY.

### F12. Validation claims and scripts (Low)

- `stage_a_precedes_final: true` is backed by `checkpoint.py:15`, which only checks that stage_a.md's mtime is earlier than final.json's. That doesn't show Stage A was saved before Stage B was *read*. The session logs do show it (see Verified).
- `final-validation.py` writes several verification flags as literals (e.g. `stage_a_precedes_final=True`, `no_future_email_leakage=True`, the test-count string). The underlying checks mostly live in `checkpoint.py`, which does assert future leakage.
- Running `final-validation.py` re-runs `checkpoint.py`, which **rewrites all six batch archives**. The Sept 28 rewrite was byte-identical, so no content changed. `validate-adjudication.py` overwrites its own record and would now fail its allowed-files assertion because of the Sept 29 handoff files. I ran neither; my own checks are listed below.
- The reviewer task prompt is stored encrypted in the parent session, so I verified what reviewers did, not the exact prompt text.

## Independently verified

- **Scope:** `git diff 540c76e..HEAD` touches only `docs/`; `git diff 540c76e -- sb corpus captures webapp` is empty. The working tree was clean. Reviewer and coordinator files were only ever added, never modified.
- **Worksheet:** removing the `agent-continuation` blocks from the current worksheet reproduces `540c76e` (SHA `cea96bae...`) apart from blank lines.
- **Regrade:** `python -m sb.regrade` gives 114/167 emails, 190 ops, 60 failing and 130 passing, matching the worksheet. The serve plan regenerated from the manifest levers is identical to `manifest.json`. The system prompt is unchanged since the capture commit.
- **Counts:** RESULTS.md has 75 unique rows: MODEL 7 (5, 11, 13, 14, 25, 30, 42), KEY 24, PLAN 8 (3, 10, 18, 20, 24, 40, 58, 60), GRADER 4 (9, 15, 36, 37), AMBIGUOUS 17, PASS 15. There are 35 queued checks. The Sept 26 split of 40 resolved (38 + 2 pilot) and 16 pending reconciles.
- **Reviewer isolation and ordering**, from Codex's native session logs (`~/.codex/sessions/2026/09/2[56]/`), all 112 reviewer sessions (96 continuation, 16 pilot):
  - Every command touched only the protocol, the reviewer's own packet and its own output directory.
  - No reviewer read `corpus/` (which contains the answer keys), `captures/`, the worksheet, coordinator files, or any other packet or review. `captures/` paths appear only as citation text inside their reports.
  - Every reviewer saved Stage A before opening any Stage B file (`stage-b.json`, after-store, trace, history), and none rewrote Stage A afterwards.
  - At most 3 reviewers ran concurrently in both phases, as approved. The continuation ran 07:14-08:22 UTC on Sept 26.
- **Hashes:** 192/192 review files, 336/336 continuation packet files and 56/56 pilot packet files match their archives.
- **Packet content:**
  - All 48 target bodies match the runner's own `render_body` output, including the 16 never read during the outage.
  - 32/32 match the observed `get_email` results.
  - Before/after stores equal the capture's day states.
  - 5,346 inbox entries across packets, none dated after the target.
  - No key or grader fields in Stage A.
- **Late delivery (Q3):** all 11 Q3 items were confirmed delivered after their requested date.
- **Tests:** `pytest sb/tests`: 107 passed (1.57 s).

## Coverage of all 75 checks

Key: ✓ = classification and rationale verified; ✓* = classification verified, rationale needs the noted qualification; ✗ = I propose a different classification; ? = needs an explicit James decision.

| Item | Current | Review | Note |
|---|---|---|---|
| 1 | KEY | ✓ | F10 authorship |
| 2 | KEY | ✓ | F10 |
| 3 | PLAN | ✓ | F10 |
| 4 | KEY | ✓ | F10 |
| 5 | MODEL | ✗ | F5: Saturday is inside "this week"; exact-day key unsupported |
| 6 | KEY | ✓ | James's own call; defensible (key pins the delivery date, model tracked the push action) |
| 7 | KEY | ✓ | F10 |
| 8 | AMBIGUOUS | ✓ | Requirement-ambiguous (F7) |
| 9 | GRADER | ✓ | F10 |
| 10 | PLAN | ✓ | |
| 11 | MODEL | ✓ | Proposal with an opt-out; default reading books it |
| 12 | KEY | ✓ | |
| 13 | MODEL | ✓* | F5: same-day delivery |
| 14 | MODEL | ? | F1 |
| 15 | GRADER | ✓ | |
| 16 | KEY | ✓ | |
| 17 | KEY | ✓ | |
| 18 | PLAN | ✓ | |
| 19 | AMBIGUOUS | ? | F1: reviewers said KEY |
| 20 | PLAN | ✓* | F11: note date slip |
| 21 | KEY | ✓ | |
| 22 | KEY | ✓* | F11 |
| 23 | KEY | ✓* | F6 |
| 24 | PLAN | ✓ | |
| 25 | MODEL | ✓* | F6: add PLAN secondary |
| 26 | KEY | ✓* | F11: three-way conflict |
| 27 | KEY | ✓ | |
| 28 | KEY | ✓ | |
| 29 | KEY | ✓* | F11: GRADER secondary |
| 30 | MODEL | ✓* | F6: Gold-tier conflict |
| 31 | KEY | ✓ | F6 root cause |
| 32 | KEY | ✓ | F6 root cause |
| 33 | AMBIGUOUS | ✓ | |
| 34 | KEY | ✓ | |
| 35 | KEY | ✓ | |
| 36 | GRADER | ✓ | |
| 37 | GRADER | ✓ | |
| 38 | AMBIGUOUS | ✓ | |
| 39 | KEY | ✓ | |
| 40 | PLAN | ✓ | |
| 41 | AMBIGUOUS | ✗ | F3: KEY |
| 42 | MODEL | ✗/? | F4 |
| 43 | AMBIGUOUS | ✗ | F3: KEY |
| 44 | AMBIGUOUS | ✓ | |
| 45 | AMBIGUOUS | ✓* | F7: infra-only |
| 46 | AMBIGUOUS | ✓* | F7 |
| 47 | KEY | ✓ | |
| 48 | KEY | ✓ | F6 root cause |
| 49 | AMBIGUOUS | ✓* | F7 |
| 50 | AMBIGUOUS | ✓* | F7 |
| 51 | AMBIGUOUS | ✓* | F7; F2 as-run false pass |
| 52 | KEY | ✓ | F6 root cause |
| 53 | KEY | ✓* | F11 |
| 54 | AMBIGUOUS | ✓* | F7 |
| 55 | AMBIGUOUS | ✗ | F8: PLAN |
| 56 | AMBIGUOUS | ✗ | F8: PLAN |
| 57 | AMBIGUOUS | ✓* | F7 |
| 58 | PLAN | ✓ | |
| 59 | AMBIGUOUS | ✓* | F7; F6 storyline |
| 60 | PLAN | ✓ | |
| B1-B14 | PASS | ✓ | Day-level only; default-time notes correct |
| B15 | PASS | ✓ | State-only pass during the outage |

Totals over all 75 checks (B1-B15 counted individually): 50 ✓, 17 ✓*, 6 ✗ (5, 41, 42, 43, 55, 56; #42 also depends on a policy question), 2 ? (14, 19). Every check was re-derived from primary sources.

## Not checked or out of scope

- I did not re-run reviewers or any model, and did not diagnose the MCP failure.
- I did not audit the ~115 unsampled passing ops. The three unsampled outage passes (F2) are in that set.
- Frontend typecheck, lint and build were not attempted; no application code was touched.
- Validators were inspected, not run (F12).
- `docs/benchmark-repair.md` still says "No independent second-review findings exist yet". I left it unchanged, per the brief not to apply corrections.

## Suggested questions for James

1. Q1/Q2 again, explicitly: keep your original #14 KEY and #19 MODEL, or adopt MODEL/AMBIGUOUS? (For #19, both reviewers said KEY.)
2. Must a multi-day trip the principal is invited to be blocked on the calendar? This decides #41/#42.
3. Re-test #5 under the fairness principle?
4. Confirm the Q4 limit ("little lead time alone is not a defect") matches what you meant.
5. Did you personally tick items 1-5, 7 and 9?

## James's responses, 2026-09-29

Recorded verbatim from this session. These settle the questions above; the corrections themselves are not yet applied to RESULTS.md, NEEDS_JAMES.md or the decision record.

1. **#14/#19:** "accept the changes." #14 MODEL and #19 AMBIGUOUS now rest on an explicit decision (F1 resolved on substance; the Sept 28 records should cite this date as the authority).
2. **Trip blocking (#41/#42):** James asked for a recommendation and whether such rules should be given to the model up front. Recommendation below; awaiting his confirmation.
3. **Re-test #5:** "sure please do." Result below.
4. **Q4 lead-time limit:** "yes leave it." No change.
5. **Items 1-5, 7, 9:** "slightly assistant drafted." Record as James's judgments with partly assistant-drafted notes (F10 resolved).

### #5 re-test under the fairness principle

- Supported by the email (served Tue Jun 9): make sure a retention conversation, which HR is putting together, happens "this week" and before the designer decides. No day, time or format is given.
- Key: an **event** on exactly **Thu Jun 11** (`serve+2d`, `exact_day`). Nothing in the text selects Thursday (the original note concedes this, K-6). Requiring an event rather than a to-do is also not established: HR organizes the conversation, and the CEO's stated role is to make sure it happens.
- Model: a to-do due Sat Jun 13 17:00, "HR wants retention conversation this week before decision is made". It fails only on kind. Its date is inside a Monday-to-Sunday week but after the working week, so whether it is timely is uncertain.
- **Result: KEY** (the exact-day and event-only requirements are unsupported), with model credit unresolved because of the Saturday due date. This matches how the audit already treats #12, #16, #34 and #35. Proposed; to be applied in the corrections pass.

### Recommendation on #41/#42 and preference rules

For the saved run, the system prompt never said whether a multi-day trip should be blocked, so do not require it. That gives #41 KEY and #42 KEY (the separate start and end markers are unsupported, and the model is not penalized), consistent with Q1 and Q9. Whether future models should block trips is a preference to state in the prompt, not something to grade silently.

### Decision, 2026-10-03

James: "score 41/42 as key." For the saved run, blocking a multi-day trip was never stated to the model, so #41 (AMBIGUOUS → KEY) and #42 (MODEL → KEY) are scored as unsupported separate start and end markers, with no model penalty. Not yet applied to RESULTS.md or the decision record; that happens in the corrections pass.

## Applied, 2026-10-03

With James's approval ("yes go ahead and include those"), the corrections pass applied F1 (authority recorded), F2 (outage amendment), F3/F4 (#41, #42, #43 KEY), F5 (#5 KEY), F6 (storyline notes and corpus-review list), F7 (AMBIGUOUS split), F8 (#55/#56 PLAN) and the F11 wording notes (#20's row was already correct, so only its worksheet note slip is recorded here), in RESULTS.md, NEEDS_JAMES.md, the worksheet (marked `independent-review` blocks), README.md, the handoff and the repair register. New records: [human-decisions-2026-10-03.json](human-decisions-2026-10-03.json) and [saved-tool-outage-amendment-2026-10-03.json](saved-tool-outage-amendment-2026-10-03.json). The September 28 decision record, original outage record, reviewer reports and batch records are unchanged. Counts: MODEL 5, KEY 28, PLAN 10, GRADER 4, AMBIGUOUS 13 (5 unclear instruction, 8 tools down), PASS 15. F9 and F12 are recorded in this review only. F10's answer is in the decision record (R7). F6's hardcoded-date pattern is in the NEEDS_JAMES corpus-review list and the register checkpoint. A read-only verification agent then re-checked scope, counts, outage figures, #43, #55/#56 and worksheet integrity and found no errors; its three wording warnings were fixed before commit.
