# SecretaryBench audit handoff for Claude Code

> **Update, 2026-10-03.** The independent review is done: [INDEPENDENT_REVIEW_2026-09-29.md](INDEPENDENT_REVIEW_2026-09-29.md). James accepted its corrections, which are applied in RESULTS.md, NEEDS_JAMES.md, [human-decisions-2026-10-03.json](human-decisions-2026-10-03.json) and the [outage amendment](saved-tool-outage-amendment-2026-10-03.json). Current counts: MODEL 5, KEY 28, PLAN 10, GRADER 4, AMBIGUOUS 13 (5 unclear instruction, 8 tools down), PASS 15. The outage spans days 51-57, not six days. Counts and outage figures below are the September 29 state, kept as history.

Updated September 29, 2026. Scope: the work completed with James on the saved Sonnet 4.5 run, the reasons for the decisions, and an independent second review by Claude Code.

## Current request: independently review the completed work

James clarified: “I just want it to verify and almost audit everything you did to give it a second look.” The next task is an independent audit of this work. The repair ideas later in this document are background, not the requested deliverable.

Treat this handoff, the coordinator's findings, and earlier validation summaries as claims to check against primary evidence. Seek mistakes and unsupported conclusions as well as corroboration. Do not assume that reviewer agreement or James's acceptance establishes factual correctness. Distinguish a policy he selected from a factual claim that still requires evidence.

Review scope:

1. Verify the checkout and inspect the changes from `540c76e` through the current HEAD, including the pilot, all six batches, human adjudication, and handoff documentation. Preserve unexpected changes. Check that source, corpus, keys, captures, grader, and scores were left unchanged.
2. Account for every one of the 75 worksheet checks. Reconstruct the supported requirement from the actual assistant instructions, email, recipient/role metadata, and available prior context before consulting that check's key or current judgment where possible. Then compare with saved tool evidence and before/after state. If conclusions have already been seen, disclose that limitation rather than claiming a blind review. Do not substitute future context or a repaired input for what the saved model received.
3. Audit the review procedure and provenance: pilot reuse, 96 continuation assignments, paired reports, Stage A/B ordering, packet integrity, input reconstruction, missing metadata, source checks, and coordinator overrides. Separate what artifacts prove from procedural claims they cannot establish. Check original human notes and reviewer outputs were preserved.
4. Challenge the substantive conclusions: KEY versus MODEL versus PLAN/GRADER, ambiguity, scope of accepted PASS results, the tool-outage evidence, and whether unavailable tools or invalid keys were wrongly converted into model credit or blame. Recompute coverage and classification counts independently.
5. Check all ten adjudication groups, including the #14/#19 amendments, against James's recorded decisions and the cited evidence. Flag any case where his words were overinterpreted or a policy preference was presented as established fact. Keep disagreements with a policy distinct from factual errors in applying it.
6. Verify the reported checks and limitations. Inspect validation scripts before running them because some rewrite historical artifacts or assume an older file scope. Use safe offline checks and source inspection; do not rely solely on a preexisting “passed” record.

Return a findings-first review with severity, affected check IDs, exact file/line or capture references, the supporting evidence, and any proposed correction. Include coverage of all 75 checks and distinguish independently verified conclusions, disagreements, and unresolved questions. If anything was only sampled or could not be checked, state precisely what and why. A clean review should explain what was independently verified, not simply repeat the existing summary.

Preserve existing results and evidence. Report proposed corrections without applying them. This authorizes review of the saved work, not corpus/grader repairs, score changes, production access, new benchmark/model runs, another reviewer campaign, or moving into the next repair phase. The independent review has not yet been performed.

## Start here

- Repository: `<repo>`.
- Working branch: `codex/phase-1d-audit-pilot`. The completed human adjudication is commit `59d02bb`; this handoff is a later documentation-only commit on the same branch.
- The saved-run hand-grade is complete for its agreed scope: all 60 failing checks and 15 sampled passing checks. Seventeen classifications intentionally remain AMBIGUOUS. That is a recorded disposition, not unfinished reviewer work.
- James's September 28 decisions supersede older pending-decision language. Read [current results](RESULTS.md), [recorded decisions](NEEDS_JAMES.md), and the [structured human decision record](human-decisions-2026-09-28.json).
- This work audited an existing run. It did not implement a future grader, repair the corpus, alter scores, or launch a new benchmark run.
- The immediate task is the independent review above. A scoped repair plan is a later possibility; implementation, a change in the approved repair order, and new paid runs are not authorized by this handoff.

## What problem we were solving

The existing grader's FAIL label did not reliably identify a model error. A failure could come from an unsupported answer key, a request delivered after its own deadline, incorrect object matching, ambiguous instructions, or unavailable tools. Loosening grading before establishing a reference could make a score rise without making the benchmark more accurate.

The immediate objective was therefore to inspect the saved evidence and separate those causes. James wanted agents to automate the manual review, with clear provenance and his involvement where judgment or policy was necessary. The resulting reference is qualified evidence, not a single unquestionable answer for every field.

The current corpus and grader already existed when this work began. In particular, the phase-2 deterministic grader was already applied. The changes described here are audit documentation, not those earlier implementation changes.

## What we did and why

| Step | Work completed | Reason |
|---|---|---|
| Restore the correct starting state | Earlier local progress was stale. The prior session refreshed `debug/model-resolution` to `540c76e`, James's `Audited to 20` commit, and preserved the 19 existing judgments. | Avoid repeating or overwriting work James had already saved. |
| Run the bounded pilot | Eight selected checks received two fresh reviews each. Six had existing human labels; two were new checks. The pilot was committed at `9a506fb`. | Check whether assisted review produced useful evidence before expanding it. Agreement was not treated as proof of correctness. |
| Approve and run the continuation | James approved 54 remaining checks across 48 emails, with two fresh reviewers per email, 96 assignments, six batches, and at most three active reviewers. | Cover the remaining worksheet while bounding usage and keeping disagreements from blocking independent cases. |
| Reuse existing evidence | Pilot #21 and B9 were reused. They were not reviewed again. | Avoid duplicate work and additional reviewer usage. The 16 pilot assignments and 96 continuation assignments were separate approved phases. |
| Determine requirements before revealing answers | Each reviewer saved Stage A before seeing the answer key, grader outcome, or evaluated actions in Stage B. | Reduce anchoring on what the benchmark or evaluated model claimed was correct. |
| Verify beyond paired agreement | The coordinator checked sources for every target and performed two additional source spot-checks per batch, selected with seed 42. | Catch shared reviewer assumptions. For example, the coordinator withheld closure on #41/#43 despite both reviewers saying MODEL. |
| Keep different questions separate | We distinguished supported requirements, stored outcomes, key defects, grader defects, and model attribution. Historical and sibling observations were recorded separately. | A bad key does not prove a good model answer; a correct day does not excuse an unsupported clock time; a reminder does not prove real-world completion. |
| Preserve the human record | Original labels and notes remain byte-preserved outside marked audit annotations. Original reviewer outputs and batch records were not rewritten. | Make revisions reviewable and retain who decided what, when, and on which evidence. |
| Record James's adjudication | On September 28, James responded to the ten grouped questions and approved recommendations for the others. | Convert the queue into explicit policy and label decisions without forcing uncertain evidence into a pass/fail result. |

### Reviewer evidence and limitations

Reviewers used fresh contexts, the same inherited model configuration, the common [protocol](REVIEWER.md), and only their assigned packet. They were not shown prior judgments or the other reviewer's findings. Isolation was procedural, not a filesystem security boundary, and the pair was not a panel of different model families.

Packets included the actual assistant instructions, target email, available earlier and same-day context, full before/after stores, and tool requests/results. They excluded future emails and model thinking. Available captured email bodies were checked against reconstruction from the unchanged corpus and saved serve plan. A reconstructed email was never treated as proof that the evaluated model read it.

Stage A omitted recipient metadata. Full observed traces retained it; the coordinator separately checked source metadata when role mattered. E27-r2 disclosed a changed ownership interpretation after reveal. For E41/#54, source metadata resolved CEO identity while the outage still prevented model attribution. E46-r1/r2 disclosed infrastructure qualifications without changing their supported-request interpretation. Original Stage A files remain intact.

## Coverage and results

| Component | Checks |
|---|---:|
| Original saved human judgments | 19 |
| Reused pilot checks, #21 and B9 | 2 |
| Fresh continuation checks | 54 |
| Total worksheet coverage | 75 |

The continuation covered 48 distinct emails with 96 fresh assignments. Several emails had multiple checks, which were assessed separately. It did not audit all passing checks in the run: 15 of 130 passes were sampled. This is not a benchmark-wide accuracy estimate.

| Current primary classification | Checks | Meaning |
|---|---:|---|
| MODEL | 7 | A supported requirement was missed. |
| KEY | 24 | The answer key contains an unsupported or incorrect requirement. |
| PLAN | 8 | Delivery timing makes the prospective request stale. |
| GRADER | 4 | Supported work was rejected by grading logic. |
| AMBIGUOUS | 17 | Required behavior or model attribution cannot be settled from the available evidence. |
| PASS | 15 | The targeted day-level or no-action check is accepted, with any other defects kept separate. |

These are primary audit classifications, not mutually exclusive root causes and not a recalculated model score. Some KEY/PLAN findings also have unresolved model credit. The 17 AMBIGUOUS classifications and the 17 infrastructure-affected checks are different, overlapping sets; their equal counts are coincidental.

At the September 26 checkpoint, 40 classifications were agent-resolved, including the two reused pilot checks, and 16 new classifications were pending. James's later responses covered ten groups, 35 queued checks, and three uncounted sibling/historical observations. They did not mean James independently re-read every underlying email. The current results distinguish saved human judgments, coordinator-verified agent findings, and application of recorded human decisions.

## James's decisions and their rationale

The complete excerpts, competing interpretations, and authority for each decision are in [NEEDS_JAMES.md](NEEDS_JAMES.md). Despite its historical filename, that file now records dispositions.

| Group | Recorded disposition | Why |
|---|---|---|
| Q1: announcements | #8 remains AMBIGUOUS; #19 changes from MODEL to AMBIGUOUS. | A fixed business date does not unambiguously require a personal calendar event under instructions permitting FYI/no action. |
| Q2: delegated choice | #14 changes from KEY to MODEL. | The email explicitly says to pick one of three valid podcast slots. Creating only a reminder to choose does not book a slot. The key already accepts multiple dates. |
| Q3: stale requests | Preserve late-delivery defects for all 11 listed checks; do not infer model failure from missing historical actions alone. | No backfill or recovery policy was supplied. Other KEY/GRADER defects may coexist. Late delivery does not identify the scheduler as the sole root cause. |
| Q4: underspecified timing | Do not require an unsupported exact deadline or hour. | “Before we announce” gives an ordering but no announcement date. “Launch night” does not select precisely 20:00. Grade supported fields; allow supported alternatives or mark a field underdetermined. Explicit duration defaults still apply. |
| Q5: conflicting context | Preserve definite key defects and flag corpus/context review; withhold strict penalties on disputed timing. | August 22 versus “next week,” subject/body timing, and conflicting rehearsal dates do not establish one reliable answer. An explicit later reschedule can be valid. An unstated dependency is not automatically a direct contradiction. |
| Q6: confirmation repair | #33 remains AMBIGUOUS under the delegated coordinator judgment. | Creating a missing, previously requested appointment is defensible, but the no-action confirmation guidance also supports leaving it alone. The invented 14:00 start remains a separate default violation. |
| Q7: conditional requests/ownership | #38 and #44 remain AMBIGUOUS, as James selected. | Neither a mandatory brainstorm task nor a general CEO oversight-task rule was supplied. |
| Q8: unavailable tools | Model-capability attribution is unassessable for all 17 affected checks. | The saved model lacked functioning secretary tools and no target-body reads were observed. State results and independent input/grader defects remain knowable. |
| Q9: insufficient specificity | #41/#43 remain AMBIGUOUS. | “Week of August 3” does not uniquely establish Monday departure or require separate trip-boundary markers. #42's missing explicit August 9 endpoint remains MODEL; #53's second-day versus +2-day key defect remains KEY. |
| Q10: approval pending | #55/#56 remain AMBIGUOUS. | Calendar changes await approval, the proposed dates are already past, and the tools failed. The false duplicate explanation for #55 is independently established, but no fulfilled move is inferred. |

James's general principle: **do not penalize the model for failing an exact requirement the input never established.** This does not make every plausible output correct or excuse violations of explicit instructions. Invalid keys, ambiguity, and infrastructure limits are not automatic model passes.

The original worksheet still contains #14's old KEY label and #19's old MODEL label as history. Their dated September 28 amendments and the current results are the effective audit judgments. Do not parse the original checkboxes alone as the current reference.

## Major discovery: the saved tool outage

[Saved outage evidence](saved-tool-outage.json) verifies six captured days: 51, 52, 53, 55, 56, and 57 (corrected 2026-10-03: day 54 also failed, so the outage is days 51-57 and 19 emails). Each has a failed secretary-server initialization, absent secretary tools, no tool calls, no observed target email retrieval, and unchanged complete stores.

This affects 16 reviewed emails and 17 checks: #45 through #60, plus B15. The grader recorded 16 FAILs and B15's no-action PASS. Those are factual state results. Neither the failures nor the incidental pass establishes how the model would perform with working tools. A successful capture wrapper did not mean that the secretary connection worked.

The audit did not diagnose the underlying connection failure, establish recoverability, change score denominators, or rerun those days. This evidence directly motivates revisiting existing repair item O-7, which concerns infrastructure errors being folded into scores.

## Why we stopped at audit documentation

- Changing the corpus during review would change the question the saved model answered. A corrected input needs its own appropriately authorized evaluation; it cannot be substituted into the old run and treated as though the model saw it.
- Changing the grader and corpus together would make a score change impossible to attribute. Repository rules require separate work and commits.
- A more permissive grader could hide real failures. #30 remains a model omission even though its key date is wrong; #42 retains a definite endpoint omission despite uncertainty about departure.
- A more semantic grader cannot invent missing ground truth. Corpus contradictions and tool failures need validity handling, not a judge guessing what the author intended.
- A new live run would spend subscription quota. It was outside the approved audit and remains subject to explicit approval.

No benchmark source, corpus, key, capture, grader, score, production system, or dependency was changed during the pilot/continuation/adjudication work. No development server was started. No repair item was automatically promoted to verified because the hand-grade finished. The application frontend was untouched.

## Proposed next step, discussed September 29

This is an earlier recommendation, not an approved implementation plan or reordered roadmap. James subsequently requested the independent review above first; do not start repair planning as a substitute for that review.

1. Prepare a scoped post-audit repair plan. Reconcile historical status text and map clear reference cases to repeatable offline acceptance checks. Keep uncertain and infrastructure-limited dimensions explicit.
2. Make **O-7 run validity** the first proposed implementation task: distinguish a failed tool connection from a valid model attempt, and investigate the connection problem before another live run. Bringing O-7 forward from phase 6 is a proposed sequencing change that needs plan approval.
3. Preserve C-1's configuration/provenance record before corpus edits. Repair delivery coherence, prose/key contradictions, and unsupported precision, under the existing K problem IDs, one ID at a time. This repairs answerability; it does not expand difficulty or rebuild the benchmark.
4. Evaluate and improve the existing grader against the reference. James favors deterministic checks for established constraints and model-assisted semantic assessment for identity and supported alternatives, with an insufficient-evidence outcome and human calibration. No hybrid grader or new grader contract has been implemented or fully specified.
5. Consider a bounded, explicitly approved live smoke test and later benchmark runs only after the relevant offline checks and validity repairs. No run is authorized by this sequence alone.

The old roadmap still has attribution/no-action work (A), configuration/reporting (C/V), corpus repairs (K), and remaining run-artifact work (O). This recommendation does not silently close or discard them. A plan must reconcile dependencies with [the repair register](../../benchmark-repair.md) and [the existing grader contract](../../grader-contract.md).

## Commit ledger

| Commit | Completed work |
|---|---|
| `540c76e` | James's existing audit through item 20, with 19 checked judgments. Starting evidence for this continuation. |
| `9a506fb` | Eight-case, 16-reviewer pilot and evidence archive. |
| `c6ad130` | Batch 1: 8 checks, 16 assignments. |
| `22be089` | Batch 2: 10 checks, 16 assignments. |
| `f98fc2e` | Batch 3: 9 checks, 16 assignments. |
| `c921c04` | Batch 4: 10 checks, 16 assignments. |
| `7ffe9ad` | Batch 5: 8 checks, 16 assignments. |
| `a9b0e53` | Batch 6: 9 checks, 16 assignments; complete integrity checks. |
| `59d02bb` | James's human adjudication, current classifications, dated amendments, and verification. |

The work was committed on the audit branch, without a push to main or production deployment. No remote publication is implied by a local commit. Existing James authorship was used without assistant attribution trailers.

## Where to find everything

Paths below are relative to this handoff unless explicitly absolute. The essential evidence is tracked; private working files are not required to understand the findings.

| Artifact | Contents and use |
|---|---|
| [RESULTS.md](RESULTS.md) | Current 75-check classifications, provenance, findings, and optional inspection sample. |
| [NEEDS_JAMES.md](NEEDS_JAMES.md) | All ten questions, original recommendations, September 28 decisions, reasons, and limits. |
| [Human decision JSON](human-decisions-2026-09-28.json) | Machine-readable authority, two label revisions, 35 affected checks, and attribution qualifications. |
| [Original worksheet](../audit-baseline-sonnet-4-5.md) | Emails, keys, model state, original human notes, agent annotations, and dated human amendments. |
| [Pilot report](../agent-audit-pilot-sonnet-4-5.md) and [archive](../agent-audit-pilot-sonnet-4-5.json) | Initial eight cases, saved interpretations and reports. Historical pending language is superseded by the September 28 decisions. |
| [REVIEWER.md](REVIEWER.md) | Continuation protocol and closure rules. |
| `reviews/E01-r1/` through `reviews/E48-r2/` | 96 original Stage A files and 96 original final reports. Not coordinator rewrites. |
| `batch-1-archive.json` through `batch-6-archive.json` | Inputs, target mappings, reviewer checksums, and references to the larger saved evidence. |
| `batch-1-coordinator.json` through `batch-6-coordinator.json` | Coordinator findings, source checks, spot-checks, and original resolution decisions. Frozen September 26 records. |
| [Outage evidence](saved-tool-outage.json) | Source lines, raw hashes, affected cases, and unchanged-state proof. |
| [Initial integrity record](final-validation.json) | Completed six-batch verification. |
| [Adjudication integrity record](adjudication-validation-2026-09-28.json) | Current counts, approved amendments, preserved history, and test results. |
| [Repair register](../../benchmark-repair.md) | Problem IDs, current dated checkpoint, older roadmap, and standing constraints. |
| [Saved capture](../../../captures/baseline-sonnet-4-5/) | Manifest, per-day stores, raw tool traces, and captured run artifacts. |

Local-only working material:

- `<repo>/.Codex/plans/codex-setup-and-handoff.md`: concise active operational handoff.
- `.Codex/plans/codex-setup-and-handoff-history-through-2026-09-28.md`: preserved earlier setup and assessment notes; historical, not current instructions.
- `.Codex/plans/audit-continuation-2026-09-25.md`: approved continuation plan and historical context.
- `.Codex/plans/audit-continuation-2026-09-26/`: immutable input packets, maps, checksums, original worksheet snapshot, and working validation scripts.
- `.Codex/plans/audit-pilot-2026-09-25/`: original pilot packets.
- `.Codex/plans/audit-adjudication-2026-09-28.md`: adjudication task record.

`CLAUDE.md` and `AGENTS.md` are local-only in this checkout. Their relevant constraints are summarized below so a fresh clone does not depend on hidden working notes. Local `CLAUDE.md` points to this tracked handoff.

### Document precedence and stale text

Use the September 28 human decision record for current adjudication. Use the untouched reviewer/batch artifacts to understand the evidence and original proposals. Older pilot and running-findings documents may say that #14/#19 await approval or that the hand-grade is unfinished; that is historical status.

The repair register and grader contract contain historical verification gates and sequencing statements that are not fully reconciled. Do not interpret one old sentence as permission to mark every grader repair verified, or treat the proposed O-7 priority as approved. Resolve that documentation in the next scoped plan.

Do not rerun `prepare.py` on the annotated worksheet. The historical `render.py` and `adjudicate.py` are guarded because regeneration would overwrite the current decision layer. Some private validators also have a fixed allowed-file scope for their original checkpoint. Preserve dated records rather than running old writers blindly over later changes.

## What was verified

- The 336 continuation packet hashes and 56 pilot packet hashes matched.
- All 192 continuation reviewer files matched archived hashes; saved Stage A preceded final output.
- Coverage matched all 75 worksheet checks, with no duplicate/missing target and no future-email leakage in reviewer inputs.
- The corpus hash, capture manifest, serve plan, and assistant prompt matched the saved evidence.
- Original human text remained intact after removing marked audit annotations. Only #14/#19 changed effective primary classification during adjudication.
- Source, corpus, capture and application files were unchanged from the starting audit state.
- Python tests: **107 passed** in the September 28 adjudication check (2.40 seconds), and again for this September 29 documentation task (2.69 seconds).
- Frontend typecheck and lint were attempted but blocked by absent dependencies (`tsc: command not found`). Build was blocked by missing Next; the vendor step was not invoked. This is not a claim that frontend checks passed.

Named verification artifacts are linked above. Source identifiers: corpus hash `03e0d963b9866d8f`; original worksheet SHA-256 `cea96bae113130e9f310cb42f66a178f3fcc49f20735fc85a1b8db34377a4e2f`. Complete manifest and packet hashes are in the archives.

September 29 handoff verification: all 29 local links in this handoff and the continuation README resolved; counts matched the adjudication record; every listed commit was an ancestor of the current branch. All 192 reviewer files matched their archived hashes, and batch records, current decisions, worksheet, and source evidence were unchanged. The documentation diff passed whitespace checks. Typecheck and lint were attempted again and remained blocked by missing `tsc`; build preflight remained blocked by missing Next, without invoking the vendor step.

Safe read-only/offline checks from the repository root:

```sh
git status --short
git branch --show-current
git log -10 --oneline
git diff 540c76e -- sb corpus captures webapp
PYTHONDONTWRITEBYTECODE=1 .venv/bin/python -m pytest sb/tests -q -p no:cacheprovider
```

The source diff should be empty at this handoff. Use the Python 3.13 virtual environment; system `python3` is 3.9 on this machine. These checks do not launch a benchmark-model run.

## Constraints for the next session

- Inspect and preserve unexpected working-tree changes. Stay on the audit branch for review; future repair work follows branch-per-problem-ID convention.
- Present a concrete plan and get approval before nontrivial implementation. One repair ID at a time; never change corpus and grader in the same commit.
- Commit completed logical tasks with James's authorship, without assistant attribution trailers. Never push to main. No push or deployment is authorized here.
- No live model/benchmark runs or extra paid reviewer panels without explicit approval. A full run is roughly 57 simulated day-turns; any later approved smoke must be bounded and run in the background.
- `sb/` is authoritative. Never edit `webapp/api/_lib/sb/`, which is generated.
- Do not touch production, `https://secretarybench.vercel.app`, the production database/corpus, or the `backups` branch.
- Do not run `scripts/recover_corpus.py` or `scripts/fix_match.py` as a reproduction step. They can fetch production or overwrite the evidence corpus.
- Any corpus repair is local, versioned, and separately scoped. Preserve old captures as evidence of the inputs actually served.
- Tests and named evidence must support any claim of verification. An accepted policy, a higher score, or reviewer agreement alone is insufficient.

## Current step and open work

Completed: the pilot, all six continuation batches, James's adjudication, and this visibility/handoff documentation.

Requested next: Claude Code independently reviews the completed work using the brief at the top of this document. No second-review findings exist yet (superseded: see the 2026-10-03 update at the top). No repair implementation or live run is in progress, and no original reviewer assignments remain.

After that review, a future task may prepare and approve the scoped repair plan, including whether O-7 moves forward in the roadmap. The unresolved scientific questions are retained as limitations: the cause of the tool outage, intended facts in inconsistent corpus passages, and requirements the supplied instructions never specified. These limits should be examined during the independent review without rerunning the original reviewer campaign.
