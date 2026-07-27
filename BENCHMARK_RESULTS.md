# SecretaryBench — multi-model benchmark results

Cross-model run of the **authored** corpus (Claude via `claude -p`, OpenAI via `codex exec`),
both driving the **same** MCP store + scheduler + grader — apples-to-apples. This file is the
durable record; scores are appended as each run finishes.

> Status: **SMOKE VERIFIED — ready for the full run.** Both providers drive the MCP tools and
> grade correctly; grading fixed (see §0). Awaiting go-ahead to run the full roster.

---

## 0. Smoke findings (2026-06-30) — two issues, one fixed, one blocking

**Smoke:** `claude-haiku-4-5`, `--limit 5`, one day-turn (5 emails).

**Issue 1 — MCP tools not reaching the model (FIXED).** The runner passed `--tools ""`, which in
the current `claude` CLI (2.1.197) means "zero tools available", so the model *narrated* tool calls
as literal text and did nothing. Fix: dropped `--tools ""` (bypassPermissions + strict-mcp-config
already scope to our MCP server) and run the subprocess in an isolated cwd (also stops `claude` from
reading the repo's `CLAUDE.md`/answer files). After the fix the model makes real MCP calls.

**Issue 2 — `match` keywords are not model-gradeable (BLOCKING).** With tools working, Haiku still
scored 1/5 — but the captured output shows it was **right**:

| email | model did | date | expected | real verdict |
|---|---|---|---|---|
| budget_meeting | created EVENT "Budget meeting for pitch comp" | Jul 5 ✓ | event Jul 5 | correct — mis-graded |
| HR_morale_sync | created EVENT "Morale discussion" | Jul 10 ✓ | event Jul 10 | correct — mis-graded |
| retreat todo | created TODO "…retreat location…notify CTO" | Jun 1 ✓ | todo Jun 1 | correct — mis-graded |
| Team_pizza_party | created EVENT "End-of-year pizza party" Jun 1 | ✗ | todo Jun 8 | genuine miss (kind+date) |

Correct email_ids, dates, and actions — but the grader ties the model's title to an obligation by
substring-matching the `match` keyword, and **138 of 140 ops (98%)** use full multi-word phrases as
`match` (e.g. `Inform Chief Technology Officer of retreat location`, `Create A List For Athlete
Visit`). No model echoes those verbatim, so correct work grades as failure. True score on this smoke
is ~4/5 (80%), not 1/5. The oracle scores 100% only because it uses the exact op name as the title.

**Consequence:** un-fixed, every model scores near-zero for phrasing reasons, not capability — the
comparison is meaningless. Fix required before the full run (see decision in chat / §6).

**Issue 2 — RESOLVED.** `scripts/fix_match.py` gives each obligation a distinctive short keyword
derived from its op name (generic scheduling words dropped; kept unique within the scenario).
**140 match sets improved; oracle stays 100%; linter passes.** 15 genuinely-ambiguous obligations
(one obligation's identity ⊆ another's) can't be auto-distinguished — they keep the original phrase
and are flagged below for a human authoring fix; they (plus a few hard cases like "notify CTO" vs
keyword `technology`) will under-count for **all** models equally, so the comparison stays fair.

**Re-smoke (`--limit 5`, same batch) after the fix — both providers work end-to-end:**

| model | driver | before fix | after fix | remaining fails |
|---|---|---|---|---|
| claude-haiku-4-5 | claude | 1/5 (20%) | **3/5 (60%)** | pizza (genuine: made event not to-do), retreat (`technology` keyword) |
| gpt-5.5 | codex | — | **3/5 (60%)** | same two; real MCP tool calls confirmed |

Flagged ambiguous obligations (kept original phrase — need a webapp authoring fix): `Company-Retreat`
{Company Retreat, Retreat Company Meeting Call}; `Marketing-campaign-new-product-delay` {Giano Ronaldo
marketing campaign, LeBron James marketing campaign}; `Partnership-with-deeptech-companies` {WHOOP HQ
Visit, WHOOP Meeting}; `Sponsoring-Marathon` {sponsorship & budget approval meeting, sponsorshippitch};
`World_Cup_Cleat_Launch` {Approve Press Embargo, Approve revised event budget, Approve tooling PO, Design
Lead 1:1, Design Lead Stage Slot}; `pizza-party` {Team_pizza_party, order-the-pizzas}.

---

## 1. Run configuration (pinned — reproducible)

| knob | value |
|---|---|
| corpus | `corpus/` (authored, recovered — see §3) |
| corpus files | 22 node files · 17 scenarios with email · **176 emails** · **140 graded ops** |
| corpus sha256 (node files) | `809d389794dd79a9…` (post match-keyword fix) |
| start date | 2026-06-01 (day 0, fixed) |
| seed | 42 |
| scheduler levers | `daily_min=1 daily_max=21 urgency_horizon=7` |
| serve span | 19 non-empty days, last mail day 18, peak 21 emails/day |
| `--days` ceiling | 30 |
| oracle (perfect model) | **176/176 = 100%** ✅ (every answer key satisfiable) |

`daily_max=21` is the **minimum feasible** for the raw authored corpus: many scenarios'
deadline-bearing emails cluster, so at the pilot default (`daily_max=5`) the schedule is
infeasible. This packs the corpus into ~19 dense days (weak retrieval-span). A later
`sb.scale` run (filler + spread) can probe span; see §6.

---

## 2. Model roster (broad — "as many as each CLI accepts")

Availability depends on the local CLIs' logins (Claude subscription; ChatGPT plan for codex);
each model's availability is confirmed at run time. Driver is inferred from the model id.

**Claude** (`--driver claude`, `claude -p`):
`claude-haiku-4-5` · `claude-sonnet-4-5` · `claude-opus-4-8` (+ `fable` if available)

**OpenAI** (`--driver codex`, `codex exec`, `model_reasoning_effort=medium`, web/image tools OFF):
`gpt-5.5` · `gpt-5.5-codex` · `gpt-5.4` · `o3` · `gpt-4o` (keep whichever the plan accepts)

Smoke test uses `claude-haiku-4-5` and `gpt-5.5` at `--limit 5`.

---

## 3. Corpus provenance & recovery (prod DB untouched)

Pulled from production `https://secretarybench.vercel.app/api/nodes` — **25 authored storylines /
184 emails**. Against `main`'s whole-day grammar, only 4 parsed as-authored, so a **local,
grading-lossless** recovery was applied (regex in `scripts/recover_corpus.py`; production DB
never modified):

- **Time-strip:** removed 80 vestigial `@HH:MM[-HH:MM]` suffixes (main grades whole-day only,
  so they carried zero graded info; answer keys unchanged).
- **Dropped 3 blank ops** (abandoned authoring stubs — empty verb name).
- **Excluded 3 genuinely broken scenarios** (need a source fix in the webapp):

| excluded scenario | reason |
|---|---|
| `rebrand-execution` | dependency cycle (`reschedule-meeting` ↔ `target-demographic`) — pick one edge direction |
| `Marketing` | 1-email stub + duplicate global anchor `@Delayed_Release_Date` with `Marketing-campaign-new-product-delay` |
| `Basketball-shoe-expansion` | 1-email stub + duplicate global anchor `@strategy_meeting` with `Partnership-with-deeptech-companies` |

Every other scenario is oracle-solvable in isolation. Result: **22 nodes / 176 emails, oracle 100%.**

Included scenarios (emails / graded-ops): `Innovation-comp` 48/17 · `World_Cup_Cleat_Launch` 22/24 ·
`Sponsoring-Marathon` 12/7 · `Partnership-with-deeptech-companies` 10/14 · `Company_Retreat` 8/6 ·
`Day-of-execution_and_Aftermath` 8/6 · `Marketing-campaign-new-product-delay` 8/7 · `pizza-party` 8/4 ·
`project_atlas` 8/7 · `Enterprise_Ai_Selection` 7/7 · `press-tour` 7/7 · `Company-Retreat` 6/11 ·
`Planning` 6/4 · `Rebrand-goes-company-wide` 6/4 · `shoe-product-launch-delays` 6/9 · `Pre-Launch` 5/6
(+ `node-5` 1/0 and 5 empty stub nodes as harmless distractors).

---

## 4. Reproduce commands

```bash
# one-time: venv + pull & recover corpus (already done)
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
PYTHONPATH=$PWD .venv/bin/python scripts/recover_corpus.py   # pull+recover -> corpus/nodes/
PYTHONPATH=$PWD .venv/bin/python scripts/fix_match.py         # model-robust match keywords

mkdir -p build

# per model — Claude:
NO_COLOR=1 ./run.sh --model claude-haiku-4-5 --driver claude \
  --corpus corpus --seed 42 --days 30 --daily-max 21 > build/run_haiku.log 2>&1

# per model — OpenAI (codex; reasoning models are slow, so a generous per-turn timeout):
NO_COLOR=1 ./run.sh --model gpt-5.5 --driver codex --reasoning medium \
  --corpus corpus --seed 42 --days 30 --daily-max 21 --timeout 600 > build/run_gpt55.log 2>&1

# optional richer report (tier × span) — MUST pass the same levers:
.venv/bin/python -m sb.analyze build/run_haiku.log --corpus corpus --seed 42 --days 30 --daily-max 21
```

The per-model `SCORE x/176 (pct)` line and per-email PASS/FAIL come straight from the runner.

---

## 5. Results

| provider | model | driver | score | pct | errored | notes |
|---|---|---|---|---|---|---|
| _(smoke)_ | claude-haiku-4-5 | claude | 3/5 | 60% | 0 | `--limit 5` verified; real tool calls |
| _(smoke)_ | gpt-5.5 | codex | 3/5 | 60% | 0 | `--limit 5` verified; real MCP tool calls |
| Claude | claude-haiku-4-5 | claude | 84/176 | 48% | 35 | completed-only 84/141=60%; 35 lost to rate window (days 11–15). [evidence](outputs/claude-haiku-4-5.md) |
| Claude | claude-sonnet-4-5 | claude | 102/176 | 58% | 0 | clean run, all completed. [evidence](outputs/claude-sonnet-4-5.md) |
| Claude | claude-opus-4-8 | claude | — | — | — | running |
| OpenAI | gpt-5.5 | codex | — | — | — | pending |
| OpenAI | gpt-5.5-codex | codex | — | — | — | pending |
| OpenAI | gpt-5.4 | codex | — | — | — | pending |
| OpenAI | o3 | codex | — | — | — | pending |
| OpenAI | gpt-4o | codex | — | — | — | pending |

_(rows filled as runs complete; unavailable models will be marked N/A)_

---

## 6. Notes / follow-ups

- **Span study (optional, later):** `sb.scale --filler 150 --days 200` buries the authored tests
  in junk to force retrieval distance — the benchmark's key discriminator — then run at default
  levers. Costs more turns; feasibility to be verified separately.
- **Source fixes for the team (webapp):** repair the 3 excluded scenarios (§3); consider a webapp
  validator change to stop authors re-adding `@HH:MM` times (the drift that broke 12 scenarios).
- **Rate limits:** runs use CLI subscription logins; the runner has exponential backoff, but a hard
  usage cap can't be outwaited. Keep an eye on long/large runs.

## 7. Code changes made for this benchmark (branch: TBD)

- `sb/live/runner.py`: pluggable **driver** (`claude` | `codex`), `--driver/--daily-min/--daily-max/
  --urgency-horizon/--reasoning/--timeout` flags. Codex specifics: `stdin=DEVNULL` (stdin-hang fix),
  session = `thread_id` for `codex exec resume`, web/image tools disabled, isolated cwd.
- `sb/analyze.py`: `--daily-min/--daily-max/--urgency-horizon` so the report's rebuilt plan matches the run.
- `scripts/recover_corpus.py`: self-contained recovery (fetches prod, applies the lossless transform,
  oracle-gates, writes `corpus/nodes/`).
- `scripts/fix_match.py`: derives model-robust `match` keywords (linter + oracle gated, idempotent).
  Final corpus sha256 `809d389794dd79a9…`.
