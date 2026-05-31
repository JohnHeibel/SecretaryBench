# Remaining Work & Known Issues

Status after the Sprint-5 remediation. **All 13 fixes in `SPRINT5_REMEDIATION.md`
are done** (calendar bootstrap, criteria date resolution, token/tool/compaction
logging, one system prompt, free-text handling, failure handling, single config,
CLI/continuity knobs, compaction dimension, stress chain, docs, scenario_id
diagnostics, response models). Full suite: **172 passing**.

This document is everything the *post-fix audit* surfaced that goes **beyond** the
original plan — mostly grading-quality issues that determine whether the *score
means what it claims*. Each item is tagged:

- 🟢 **CODE-FIXABLE** — we can fix it without touching the dataset.
- 🟡 **DATASET-BOUND** — the real fix is in `Emails.xlsx`, which is the frozen
  source of truth. Code can only mitigate.
- ⚪ **OUT OF SCOPE / BY DESIGN** — intentional, or a larger separate effort.

**Decided approach (supersedes the earlier "Excel frozen" stance):** Miguel
edits `Emails.xlsx` success criteria **by hand** to fix wrong/missing dates; the
**code** must faithfully grade whatever the criterion specifies. No code
workarounds that paper over bad criteria. Only Miguel edits the sheet — code
never auto-edits it. See `HANDOFF.md` §4 for the exact next-actions sequence.

---

## A. Grading correctness (these change what the score means)

### G1. Bare `CC-{date}` false-negatives — RESOLVED BY DECISION (split into code + manual)
**~8–10 of 30 calendar scenarios mark a *correct* model wrong.** The sheet uses
bare `CC-{date}` when the email targets a date the token system can't express
("3rd Friday of November", "first Wednesday of August", "next Tuesday"). FIX-2
resolves bare `{date}` to the arbitrary *served day* and strict-matches it, so a
model that schedules on the real date fails.
Confirmed by a 15-judge fan-out: false-negatives = **C19, S32, C13, C21, C11,
C20, C22, C23** (plus likely C12, S10).
**Decided fix (NOT a leniency hack):**
- bare `{date}` stays **strict** (= served day);
- where the proper date is **already a token** the resolver can't read (Group 2 —
  `{nextweek-wednesday}`, `{third Friday}`, etc.), fix the **resolver** (G5) →
  grades correctly, no Excel edit;
- where the criterion is genuinely the wrong token (bare `{date}` but the email
  means another date), **Miguel relabels it in Excel** after G5 gives him tokens.
See `HANDOFF.md` §4 for the exact sequence (resolver → TC check → re-run → short
manual list).

### G2. TC (todo) due-date is never checked — 🟢 CODE-FIXABLE  `[HIGH]`
The grader's TC branch only checks a todo **exists** (+ count, + literal content
token). It never compares the todo's `due_date`. Of 48 TC sub-criteria, **21
specify a concrete deadline** (`TC-{date-nextweek}`, `TC-{nextweek-date +4}`, …)
that is currently **ignored** — a model can set any deadline and pass.
**Fix:** add due-date matching to the TC branch, mirroring CC (`_event_matches_date`),
gated by the same bare-`{date}`-is-lenient rule from G1. Fills the hole for the
specific-deadline todos; bare `TC-{date}` stays existence-only.
**Note:** some specific TC tokens (`{nextweek-wednesday}`, `{nextweek-friday}`)
aren't handled by the resolver (see G5), so they'd stay existence-only until G5.

### G3. Email replies (`send_email`) are never graded — 🟡 DATASET-BOUND  `[MED]`
The criteria vocabulary is only `TC`/`CC`/`RS`/`No action`. There is **no prefix
for "a reply was sent"**, so a model's email replies are invisible to scoring.
Many scenarios that should reward a reply only grade No-action or a todo.
**Code can't fix this alone** — it needs a criterion type in the dataset (e.g.
`EM-<recipient>`). Mitigation: we *log* every `send_email` in `tool_calls.jsonl`,
so reply behavior is observable even though it's unscored.

### G4. Three free-text "action" criteria are ungraded — 🟡 DATASET-BOUND / 🟢 partial
`delete meeting {date-14th} {date-11AM}`, `Remove meeting on {date-1:14PM}`,
`create new meeting on {date-3PM}` describe real actions but carry no `TC`/`CC`/`RS`
prefix, so FIX-5 reports them as `ungraded` (honest, but unscored).
**Code mitigation:** teach `grader` to recognize the verbs (delete/remove/create
meeting → an event delete / `CC` check). Doable but brittle (NLP on 3 strings).
The clean fix is a dataset relabel (frozen). Recommend leaving ungraded unless
these scenarios matter — it's already better than the old silent auto-pass.

### G5. Date-token resolver gaps — 🟢 CODE-FIXABLE (partial)  `[MED]`
`engine._resolve_one_token` doesn't handle several forms the sheet uses, so those
criteria stay unresolved → existence-only (no date check):
- weekday tokens: `{nextweek-wednesday}`, `{nextweek-friday}`, `{date-Tuesday}`
- ranges: `{date-12:30-2:00PM}`
- prose dates the sheet writes in the **body** but not as a token: "3rd Friday of
  November", "first Wednesday of August", "two weeks from this Monday".
**Fix the tractable ones in code:** add "next-week + weekday" and "Nth weekday of
month" patterns to the resolver. The prose-only ones (G1's hard cases) can't be
graded strictly without a token in the *criterion*, which is dataset-bound.

### G6. Content matching is substring-based — ⚪ BY DESIGN  `[LOW]`
`TC-item3` passes if a todo title/description *contains* "item3". Brittle but
adequate at benchmark scale. Leave as-is unless false matches show up.

### G7. "No action" override is heuristic — ⚪ BY DESIGN  `[LOW]`
`engine._grade_email_against_diff` flips a passed "No action" sub-criterion to
fail when the model took *any* action (incl. a delete invisible in the diff).
Works; documented. Watch for edge cases but no change needed now.

---

## B. Coverage & harness

### H1. Only one harness actually runs — ⚪ OUT OF SCOPE  `[plan: optional FIX-14]`
`claude -p` works; `CodexAdapter` is a stub. The "any harness" claim is a clean
*interface*, not a proven second implementation. Proving it (Codex CLI + the
`BENCH_MODE` server-side tool gate, §7.5/§7.6) is a separate effort.

### H2. Compaction detection is a heuristic — ⚪ BY DESIGN  `[LOW]`
Claude Code's print mode compacts **silently** (no `system/compaction` event on
2.1.158), so we detect it by a large context-size drop between turns. Verified
live (196K→18K). It's inference, not an explicit signal — fine, but note it.

### H3. `context_window_exceeded` is always `False` — 🟢 CODE-FIXABLE  `[LOW]`
The field exists but is never set true (claude compacts before overflowing). If
you want a real "exceeded" signal, set it when a turn errors on context length.
Low value today.

### H4. OpenRouter path is wired but unverified — ⚪ BY DESIGN  `[plan: D4]`
`--openrouter` / `--api-base` thread `ANTHROPIC_BASE_URL` + key into the adapter,
but no run has confirmed a non-Anthropic model completes a scenario. Left as-is
per the plan (D4). Verify before relying on cross-provider scores.

---

## C. Verification gaps (not bugs — things we haven't *proven*)

### V1. No full 100-day / 109-scenario live run — `[HIGH to close]`
We validated on a 10-day / 12-scenario live slice + small runs + 172 tests. The
whole dataset has never run end-to-end live. Nothing suggests it won't, but it's
the only thing that proves the full run is clean and gives a real headline score.

### V2. Test hygiene — 🟢 CODE-FIXABLE  `[LOW]`
`tests/test_pipeline.py` is a homegrown harness that runs simulations at *import*
(~35s, needs the server). It works but is slow and non-idiomatic. Could be
pytest-native with fixtures. Also: `pytest-timeout` isn't installed, so a hung
live call can't be auto-killed in CI.

---

## D. Confirmed fine — do NOT "fix" these (adversarially checked)

- stream-json **message-id dedup** is correct (counts usage once).
- `resume_session` no-op is intentional (resume happens in `run_turn` via `--resume`).
- day-100 offset **clamping** holds; the small `remaining_active` tail is acceptable.
- the single source of truth (`harness/base.py`) — one prompt, one MCP config.

---

## Order to execute (see HANDOFF.md §4 for detail)

1. **CODE — G5 resolver tokens** (weekday `{date-Tuesday}`/`{nextweek-wednesday}`,
   Nth-weekday `{third Friday}`, date+time combos). Recovers Group 2 with ZERO
   Excel edits — the proper date is already in those criteria.
2. **CODE — G2 TC due-date matching** (mirror CC; bare/unresolvable → existence-only).
3. **CODE — `docs/TOKEN_REFERENCE.md`** so Miguel's hand edits use valid tokens.
4. **RUN** the 10-day slice / full run; find scenarios that STILL score wrong.
5. **MIGUEL (by hand, in Excel)** — relabel only the still-broken criteria, using
   the short exact list produced in step 4. Known trivial: S21
   `TC-{greenlight product A}` → `TC-greenlight`.

Genuinely dataset-bound (only Miguel can fix, code can't invent data): **G3**
(no email-reply criterion type), **G4** (3 free-text action criteria), and the
underspecified scenarios (`{deadline}` flag-ambiguity, `{C}` abstract date A/B/C).
Keep Miguel's manual surface small — do all the code and the re-run first.
