# Continuation of the saved Sonnet 4.5 audit

**Start with the [Claude Code handoff](CLAUDE_CODE_HANDOFF.md)** for the complete work history, reasons, current decisions, evidence map, and proposed next step. Updated September 29, 2026.

<!-- human-adjudication:start -->
**Current: September 28 human adjudication.** James’s responses cover all ten decision groups. #14 is now MODEL and #19 AMBIGUOUS; their original worksheet labels remain as history. Other underspecified cases retain AMBIGUOUS or qualified model attribution. [Current results](RESULTS.md), [recorded dispositions](NEEDS_JAMES.md) and [decision record](human-decisions-2026-09-28.json) supersede the historical pending language below. All original reviewer and batch records are unchanged. The following methodology and counts describe the September 26 audit checkpoint.
<!-- human-adjudication:end -->

Authorized by James on 2026-09-26. Starting branch `codex/phase-1d-audit-pilot`, commit `9a506fb`; clean tracked checkout. Six approved batches cover 54 checks across 48 emails, with two fresh reviewers per email and at most three concurrent. This is a one-time assisted manual audit, not a future benchmark grader.

The existing 19 human judgments and their notes are preserved. Pilot #21 and B9 are reused without new reviewer calls. #14 and #19 remain human decisions; settled defects in #16 and #20 do not imply model credit.

Each reviewer receives only the [common protocol](REVIEWER.md), one packet, and its output location. Fresh contexts have no forked conversation or model override. Stage A obligations are saved before Stage B reveals the key, grader outcome and evaluated actions. Isolation is procedural, not an operating-system access boundary. Reports in `reviews/` are original reviewer outputs, not coordinator conclusions.

The coordinator checks source evidence and identities for every target. Paired agreement alone does not close a case. `AGENT_RESOLVED` identifies a supported classification; `NEEDS_JAMES` identifies an unresolved classification or proposed revision to a saved human label. A settled KEY/PLAN defect may retain a separate unresolved model-credit question. Day-level target outcomes remain distinct from clock-time/default defects and sibling or historical observations.

Packets reconstruct available inbox text from the unchanged corpus and exact captured serve plan. Where a captured get_email result exists, its body is checked against the reconstruction. Missing read evidence is explicitly identified and available infrastructure status supplied. Full day stores and raw tool requests/results are included; model thinking, future emails, prior judgments and sibling reviews are excluded. Batch archives preserve Stage A/B inputs and packet checksums, referring large stores/traces to the existing capture.

Coverage and findings are recorded at each batch checkpoint. A deterministic additional source spot-check uses seed 42, one otherwise resolvable failed check and one passing check per batch. These are coordinator checks, not a third independent vote.

Verification baseline: existing Python tests passed, 107/107 in 2.77 seconds. Webapp typecheck and lint could not run because dependencies are absent (`tsc: command not found`). Build is blocked by the same missing dependencies; its vendor step is not run because this task must not alter the generated application copy. No application/UI source changes or server startup are included. No new benchmark runs or score recalculation are performed. Repair statuses are not promoted.

The [saved tool outage evidence](saved-tool-outage.json) covers 16 emails and 17 targeted checks. A successful capture wrapper did not mean functioning secretary tools. Disconnected days have no observed target-body retrieval, so state failure and model-quality attribution remain separate.

Coordinator disagreement can withhold closure even when reviewers agree: batch 4 leaves #41/#43 pending after expanded source inspection of the trip wording. E27-r2 disclosed a changed ownership interpretation after recipient metadata appeared in the captured result; the original Stage A is retained and #44 is pending. Uncounted sibling questions remain in the decision queue without increasing coverage.

Stage A reconstructs sender, subject, body and delivery date, but omits recipient metadata. Full observed traces retain it. For E41 (#54), the coordinator checked the original corpus recipient and role declaration to resolve the subsidiary CEO/CFO identity question; the outage still prevents model attribution. E46-r1/r2 flag changed evaluation after Stage B supplied infrastructure evidence, while explicitly retaining their original supported-request interpretation and Stage A files. These disclosures are not silently normalized.

Completed all six batches: 54 fresh checks across 48 emails, 96 initial assignments, two reused pilot checks and 19 preserved human judgments account for all 75 worksheet entries. Forty agent classifications are resolved and 16 remain pending. The ten decision groups cover 35 checks with interpretation, model-credit or saved-label questions, plus three uncounted observations. [Results](RESULTS.md), [decision shortcuts](NEEDS_JAMES.md) and [final integrity record](final-validation.json) provide the completion checkpoint.
