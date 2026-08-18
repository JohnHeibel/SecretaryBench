# SecretaryBench repair register

The single living record of what is wrong with this benchmark, what we changed, and
what has actually been proven. Read this before starting work; update it in the same
turn as any code change.

**Status legend:** `open` → `investigating` → `fix proposed` → `applied` → `verified`.
Nothing reaches `verified` without a named artifact (a passing test, a re-scored run,
a diff). IDs are permanent and never reused.

---

## ⚠ Read this before running anything

Four things in this repo will destroy evidence or produce a confidently wrong number.
None of them announce themselves.

1. **`scripts/recover_corpus.py` and `scripts/fix_match.py` are documented as the
   "reproduce" path (`BENCHMARK_RESULTS.md:146-147`) and are not.** `recover_corpus.py:25,32`
   fetches `https://secretarybench.vercel.app/api/nodes` **live from production** — which
   `CLAUDE.md` lists as do-not-touch — and `fix_match.py:21` rewrites `corpus/nodes/` **in
   place** (`recover_corpus.py:131,133`, `fix_match.py:86,88` unlink then rewrite). Running
   either would overwrite the corpus the three surviving runs were scored against, which is
   what makes their levers recoverable at all (C-1). **Do not run either script**; the
   production fetch alone violates a standing rule.

   *Corrected by the verifier pass:* an earlier draft of this box said the damage was
   **permanent**. It is not. Every file both scripts touch is git-tracked, committed and
   clean (`git ls-files corpus/` → 16 files; `git status --porcelain corpus/` → empty;
   unchanged since `24331fb`), so `git checkout -- corpus/nodes` restores them byte-for-byte
   and the lever recovery is reproducible from git at any future date. The hazard is real
   and the recovery is not perishable. This correction is why phase 1 is no longer labelled
   time-critical — see the phase table.
2. **`BENCHMARK_RESULTS.md:89-91` says `daily_max=5` is infeasible for the authored
   corpus. That is false for the corpus in the repo** — `daily_max=5` is feasible at
   `--days ≥ 57`, `daily_max=4` at `≥ 68`; only `daily_max=3` is infeasible at any ceiling
   (C-9). A reader picking levers for a paid run from that sentence picks the wrong one.
3. **`sb.analyze` silently reports a different span distribution if you pass the wrong
   levers**, and its defaults disagree with the runner's (C-4). It never reads the levers
   out of the log it is analysing. There is no error path; you get a plausible graph.
4. **Do not assume the plan builds.** 19 of 100 seeds raise `InfeasibleSchedule` on the
   current corpus, and `daily_max=3` fails at any `--days` (K-2). Seed-variance robustness
   checks cannot be run until that is fixed.

Standing rules from `CLAUDE.md` that phase A did not change: no live model runs without
explicit approval; never move the grader and the corpus in one commit; `sb/` is
authoritative and the vendored webapp copy is generated; production DB, the deployed
webapp and the `backups` branch are off limits.

---

## Start here (next session)

Phase 0 and phase A are complete and committed, and the revised phase table is signed off.
**The next action is phase 1: freeze the record before anything writes to `corpus/`** —
capture the recovered levers, the plan digests, and the ~270 model-authored titles
harvested from the four logs into a committed sidecar. It is free, offline, and needs no
live run.

Two standing consequences of the sign-off, carried here so they are not re-litigated:
G-1, G-3 and G-8 may reach `fix proposed` in phase 2 but cannot reach `verified` before
phase 7; and V-1's rescope-or-rebuild decision belongs to **phase 8**, after the run.

**The scope rule for phases 1 through 7, in one line:** fix what is broken about the
benchmark we have. Do not make it harder. What a repaired-but-trivial score implies is
phase 8's problem.

**Revised 2026-08-18: a baseline run moves to the front (phases 1a-1c).** The O-timing
decision taken at sign-off assumed no run would happen before phase 7, which made O's
payoff land too late to matter. That assumption no longer holds. Building the minimum
capture slice *before* the run converts one paid run into a permanent, re-gradeable asset:
every later grader change can be scored against it offline, for free, indefinitely. It
also produces the first artifact in this project's history whose served model is certified
rather than asserted (see the provenance note under C-1).

Phase A ran six read-only category agents over the shared brief
(`docs/benchmark-repair-evidence.md`). They produced **50 findings**. Their full working
— every command, every table, every quoted log line — is on disk and is *not* reproduced
here:

| section | path | findings |
|---|---|---|
| G grader identity | `docs/_repair/G.md` | 10 |
| A attribution / no-action | `docs/_repair/A.md` | 6 |
| O run artifacts | `docs/_repair/O.md` | 9 |
| C config / provenance | `docs/_repair/C.md` | 9 |
| K corpus authoring | `docs/_repair/K.md` | 8 |
| V construct validity | `docs/_repair/V.md` | 8 |

This register carries each finding's problem, severity, single best piece of sourced
evidence, and options. **Go to the section file for the working.**

A verifier pass has red-teamed this merge against the code; its verdict is
`docs/_repair/VERIFY.md`. It checked ~200 claims: ~180 confirmed, 2 substantively wrong,
8 overstated, 3 unfalsifiable, and all six headline numbers reproduce to the digit. Its
corrections are applied here and each is marked in place. Two are worth knowing before
you read anything else: the corpus hazard is **not** permanent (`corpus/` is git-tracked
and restorable), and O-1's "nothing can ever be re-graded" is too strong. Both fed the
phase table, which has been re-argued on G-4 alone.

**The revised phase table was signed off on 2026-08-17, as written below.** Two decisions
were taken at sign-off:

1. **O stays at phase 6.** The consequence was put explicitly and accepted: G-1, G-3 and
   G-8 can be *designed* in phase 2 but cannot be validated until the phase 7 run, because
   of O-1's corrected limit. The alternative — one extra instrumented run up front to
   unblock them — was declined as not worth the cost.
2. **V-1's direction is decided at phase 8, after the paid run, not before it.**
   *(Revised 2026-08-18. The sign-off originally parked this at phase 4; it has been moved
   to the end.)* Phase 5 is therefore a corpus **repair**, unambiguously. The repair fixes
   what is broken about the benchmark we already have; it does not try to make the task
   harder. If the repaired benchmark turns out to be trivial for a frontier model, **that
   is itself a result and worth reporting**, and it is a better basis for deciding what to
   build next than a forecast is. Nothing about scaling, rescoping the retrieval claim, or
   re-authoring the tier set is in scope before phase 7 reports.

### Why the phases are ordered the way they are (revised after phase A)

The pre-phase-A argument was: *the store state is thrown away when a run ends, so the only
way to evaluate a grader change is to pay for a whole new run; therefore do **O** first so
every later grader change can be re-scored offline against runs already paid for.*

**One finding falsifies it. A second, which an earlier draft leaned on equally, does not
survive the verifier pass — the honest version of both is below.**

- **G's central result was produced offline with no live run at all.** The title-policy
  sweep (G-4) ran the *shipped* `sb.engine` + `sb.grader` against the *shipped* corpus and
  showed a scheduling-perfect agent scoring 92/167 = 55%. G already has a working offline
  evaluation loop. It never needed O. **This is the whole load-bearing argument for the
  re-ordering**, and it was independently reproduced by the verifier.
- **O-1, in its defensible form.** The store is terminated at `runner.py:513-516` with
  nothing read out of it, and every surviving artifact is hand-copied stdout. An earlier
  draft concluded "no run already paid for can *ever* be re-graded". That is overstated,
  and the register contradicts it three times: G-2 re-grades opus 90→97, A-1 re-grades
  90→55 day-scoped, A-6 re-grades 90→91 and 90→89. The verifier reproduced two of those
  from the logs alone. The true limit is narrower and still material: **no paid run can be
  re-graded under a rule that would newly admit objects the log never rendered** — anything
  that widens `_title_hit`, changes `op.kind`, or re-attributes by `email_id` — because
  `grader.py:151-152` filtered those out before printing. So G-1, G-3 and G-8 are blocked
  from offline re-scoring; rules that relax the count check, re-scope the no-action delta,
  or drop an object are evaluable today. **This weaker version does not by itself move O
  from phase 1 to phase 6**; G-4 does.

One further finding forces a move:

- **C-1 argues for capture early, but not under a deadline.** The levers of the three
  surviving runs are recoverable by replaying their printed serve plans against the corpus
  they were scored on. `BENCHMARK_RESULTS.md` documents a command that overwrites that
  corpus in one step. Capturing first is free and beats re-deriving later — but `corpus/`
  is git-tracked and clean, so the recovery is reproducible from git regardless. **Do it
  early; it is not perishable.**
- **V-1 decides what the corpus pass is for.** The authored corpus is ~10.9k tokens and
  the largest needle gap is ~5.4k tokens of mail against a ≥200k window: no authored fact has ever
  scrolled out of context. Either the corpus gets scaled until it does (V-1 opt 1, gated on
  V-5) or the retrieval claim gets rescoped (V-1 opt 2). That decision determines whether
  K's pass is a repair or a rebuild, so it must precede K, not follow it.

Revised table. Old phase numbers in brackets.

| phase | category | why here | cost |
|---|---|---|---|
| 0 | M, E | done — the reported blocker | free |
| A | — | the register (done) | free |
| 1 | **C-1** *(new)* | a documented command overwrites the corpus the three runs were scored against. Capture the recovered levers, the plan digests, and the ~270 model-authored titles harvested from the four logs before anything writes to `corpus/`. Recoverable from git if missed, so early rather than urgent | free, do it early |
| **1a** | **O-1, O-3, O-5a** *(new)* | the minimum re-gradeable capture slice: `--out DIR`, an end-of-run store dump, and **recording objects that matched no keyword**. Built before any run so the run is a permanent asset rather than a fourth disposable log | free |
| **1b** | — *(new)* | bounded smoke (`--limit`, background) to prove the capture writes what phase 2 needs, and that the runner certifies the served model | cheap, bounded |
| **1c** | — *(new)* | **the baseline run**: certified model, known config, fully re-gradeable. Run at seed 42 / `daily_max=5` so it is directly comparable to `outputs/opus.md` and `outputs/sonnet.md`, which retroactively validates them | **paid** |
| 2 | G *[was 2]* | the grader. Now testable against **real recorded behaviour** from 1c, not only the oracle title-policy sweep | free to iterate |
| 3 | A *[was 3]* | depends on the identity contract G settles; A's no-action rule is 56–57% of every recorded score | free |
| 4 | V *(part)*, C *[was 5]* | pin one lever and one config (C-2, C-3, K-2) and fix the **reporting** gaps: report V-3's null floor beside every score, make V-6's by-tier report exist, stamp provenance. **Does not decide what the benchmark claims** — see phase 8 | free |
| 5 | K *[was 4]* | the corpus pass, at the pinned lever, against G's keyword contract. **A repair, not a rebuild**: make every authored question answerable and every keyword discoverable. Do not add difficulty, do not scale, do not re-author the tier set | free |
| 6 | O *(remainder)* *[was 1]* | the O items **not** in the 1a slice: O-2 trace loss, O-6 to O-9. O-4 (retrieval observability) is here rather than 1a because it requires changing the model-facing tool surface, a benchmark decision that must be taken with G/A/K/V settled | free |
| 7 | — | the full instrumented run of the **roster** at the pinned config, plus the hand-grade. 1c is one model as a baseline; this is the comparison | **paid** |
| 8 | **V-1, V-5** *(new)* | reassess with a real number in hand. Is the repaired benchmark still measuring anything interesting, and does the retrieval claim survive? A trivial score is itself a result. **Nothing here is scoped until phase 7 reports** | free to decide |

**Phase 1.5 no longer exists as scheduled, and this is the honest version of why.** The
old table listed "hand-grade ~30 emails: the honesty baseline" at cost `free`. O-5
establishes it is not free and not possible: `grader.py:152,168` renders only objects that
already matched an answer-key keyword, and the store was discarded, so no surviving
artifact shows what a model actually did on the ~85% of failures that read
`(nothing matching created)`. A human cannot hand-grade behaviour no artifact records. It
splits in two:

- **1.5a — audit the answer keys against the email prose.** Genuinely free, and *K already
  did most of it*: 18 emails delivered after a date their own body states (K-1), 8 emails
  whose prose contradicts their answer key (K-3), 5 exact-day answers with no cue in the
  body (K-6), 10 rendering defects including one leaked answer-key expression (K-5).
- **1.5b — hand-grade real model behaviour.** Needs O plus a fresh instrumented run.
  It is part of **phase 7**, not a precursor to phase 2.

The consequence for phase 2's honesty is stated plainly: **a G change can be evaluated
offline but cannot be verified as "toward the truth" until phase 7.** Under the status
legend that means G fixes may reach `applied` and must not reach `verified` before then.
The partial substitutes available now are the oracle title-policy sweep (G-4) and the
harvested-title replay (G-4 opt 2), both free, both biased toward objects that already
matched something.

---

## Dashboard

| ID | problem | severity | status | phase |
|---|---|---|---|---|
| E-1 | venv built on macOS system Python 3.9; `mcp` and `mcp_app.py` need 3.10+ | blocks-running | **verified** | 0 |
| E-2 | `mcp>=1.0.0` unbounded resolves to 2.0.0, which moved `FastMCP`; tool server never starts | blocks-running | **verified** | 0 |
| E-3 | `start_mcp` discarded stderr, so any startup crash read as "did not come up" | slows-work | **verified** | 0 |
| M-1 | `run.sh` `live` branch dropped every CLI argument, silently running defaults | blocks-measurement | **verified** | 0 |
| M-2 | runner never reported which model the CLI actually served | blocks-measurement | **verified** | 0 |
| M-3 | `--model` behaviour across `--resume` turns unverified on CLI 2.1.233 | distorts-measurement | **verified** | 0 |
| M-4 | documented model roster may be stale | cosmetic | **closed** | 0 |
| M-5 | runner inherits `CLAUDE_CODE_*` env when launched from a Claude Code session | unknown | open | 0 |
| **G-1** | required `match` keyword absent from the mail the model reads (32/125 ops) | blocks-measurement | open | 2 |
| **G-2** | exactly-one rule counts keyword hits, not obligations; 0 of 57 "duplicates" are duplicates | blocks-measurement | open | 2 |
| **G-3** | `description` is in the match haystack, invisible in the log, unmentioned to the model | blocks-measurement | open | 2 |
| **G-4** | the oracle certifies satisfiability and is structurally blind to gradeability | blocks-measurement | open | 2 |
| **G-5** | lint #5 compares author strings to author strings; name-aware variant flags 10 | distorts-measurement | open | 2 |
| **G-6** | pool is cumulative over the run, so obligations get harder by position (52%→26%) | distorts-measurement | open | 2 |
| **G-7** | `cancel` graded as keyword absence over the cumulative pool | distorts-measurement | open | 2 |
| **G-8** | `match` defaults to the whole obligation name as one contiguous phrase (21 ops, 13% pass) | distorts-measurement | open | 2 |
| **G-9** | the grader's own reason strings cannot distinguish its failure modes | slows-work | open | 2 |
| **G-10** | the identity logic has no unit tests | slows-work | open | 2 |
| **A-1** | 56–57% of every score is the abstain check; the model ranking is not robust to the rule | blocks-measurement | open | 3 |
| **A-2** | the live attribution split is untested and diverges from the path the oracle validates | blocks-measurement | open | 3 |
| **A-3** | five ways over-action escapes the no-action check; only one is the one §2.6 names | distorts-measurement | open | 3 |
| **A-4** | `_watch_attribution` is blind to the sibling stamp it is cited as evidence about | distorts-measurement | open | 3 |
| **A-5** | a wrong `email_id` silently removes the object from every pool; honest model loses credit | distorts-measurement | open | 3 |
| **A-6** | worked case: one stale stamp caused one false PASS and cost one earned PASS | distorts-measurement | open | 3 |
| **O-1** | the harness writes nothing; surviving artifacts are hand-copied stdout | blocks-measurement | **applied** | **1a** |
| **O-2** | tool trace drops calls by message-id collapse; loss is model-dependent | blocks-measurement | open | 6 |
| **O-3** | a final-state dump is insufficient: the store records no history and no day | blocks-measurement | **applied** | **1a** |
| **O-4** | retrieval is unobservable server-side; only the lossy client trace sees it | blocks-measurement | open | 6 |
| **O-5** | the log renders only keyword-matched objects, so the dominant failure is unfalsifiable | blocks-measurement | **applied** *(capture half; the log still renders only matches)* | **1a** |
| **O-6** | tools and narration attributed per day not per email; narration truncated at 200 chars | distorts-measurement | open | 6 |
| **O-7** | infra errors and retries fold into the score with no machine-readable marker | distorts-measurement | open | 6 |
| **O-8** | no cost, timing, token or version capture, although the CLI hands them over | slows-work | open | 6 |
| **O-9** | objects with an unparseable date are dropped from grading without a trace | distorts-measurement | open | 6 |
| **C-1** | no artifact records its levers; two documented commands overwrite the corpus they are recovered from | blocks-measurement | open | **1** |
| **C-1b** | the model label on every surviving artifact is asserted, not observed (M-1/M-2 were live) | blocks-measurement | open | **1** |
| **C-2** | `urgency_horizon` absent from the score stamp and moves 148/167 serve dates | distorts-measurement | open | 4 |
| **C-3** | corpus satisfiability is lever-dependent; the oracle gate is hardcoded to one setting | distorts-measurement | open | 4 |
| **C-4** | `sb.analyze` takes levers by hand and silently reports a different span grid if wrong | distorts-measurement | open | 4 |
| **C-5** | corpus identity asserted by a hash whose algorithm exists nowhere in the repo | distorts-measurement | open | 4 |
| **C-6** | the whole evidentiary base entered in one commit titled "." that also rewrote the harness | distorts-measurement | open | 4 |
| **C-7** | the stamp prints "served <model>" from the requested model when nothing was served | distorts-measurement | open | 4 |
| **C-8** | run logs default to a gitignored path; artifacts preserved only by hand | slows-work | open | 4 |
| **C-9** | residual stale claims in the durable record the phase-A1 banner does not cover | slows-work | open | 4 |
| **K-1** | the serve plan drifts past the dates the corpus narrates (18 emails, 11 ops) | blocks-measurement | open | 5 |
| **K-2** | corpus date-coherence needs a high `daily_max`; the two headline runs used 5 | blocks-measurement | open | 5 |
| **K-3** | prose states a date the answer key contradicts (8 emails) | distorts-measurement | open | 5 |
| **K-4** | `Innovation-comp` supplies 29% of emails and ~38% of every model's score | distorts-measurement | open | 5 |
| **K-5** | rendered tokens fuse into prose, render at the wrong grain, or leak authoring notes | distorts-measurement | open | 5 |
| **K-6** | answers the prose does not pin: 5 no-cue exact days, 12 `by:` windows averaging 19.5d | distorts-measurement | open | 5 |
| **K-7** | four `kind` choices contradict the obligation; 14 `eq` ops land on a weekend | distorts-measurement | open | 5 |
| **K-8** | anchor, name and metadata hygiene: 11 unused anchors, a case-pair, 9 whitespace names | slows-work | open | 5 |
| **V-1** | the corpus is two orders of magnitude too small to push any fact out of context | blocks-measurement | open | **8** |
| **V-2** | the search-rate figure is real for opus and unmeasurable for sonnet and haiku | blocks-measurement | open | 6 |
| **V-3** | 38.3% of the score is available to a model that never calls a tool | distorts-measurement | open | 4 |
| **V-4** | 190 op-level judgements compressed into 167 binary points | distorts-measurement | open | 4 |
| **V-5** | `sb/scale.py` fails non-monotonically and raises the do-nothing floor to 60–68% | blocks-measurement | open | **8** |
| **V-6** | `tier` is loaded and never read; the by-tier report hangs off a file nothing writes | blocks-measurement | open | 4 |
| **V-7** | the authored tier gradient does not exist on the axis it is defined by | distorts-measurement | open | 4 |
| **V-8** | the span axis is reconstructed from levers the artifact never recorded | distorts-measurement | open | 4→6 |

**Counts.** 50 phase-A findings: 18 blocks-measurement, 26 distorts-measurement,
6 slows-work. Cost to verify: 49 free-offline, 1 needs-one-live-run (A-5). All 50 are
`open`; nothing has been applied.

---

## Phase 0 — model resolution and environment (complete)

### E-1 · venv on the wrong Python
**Status:** verified.
`/usr/bin/python3` on this machine is 3.9.6. `mcp` requires 3.10+, and
`sb/live/mcp_app.py:50,81` use PEP 604 (`str | None`) in FastMCP tool signatures, which
FastMCP evaluates at runtime — a hard TypeError on 3.9. The documented setup line
(`run.sh:18` @`24331fb`, `BENCHMARK_RESULTS.md` §4) says bare `python3`, which builds a 3.9 venv.
**Fix:** build with an explicit interpreter — `uv venv --python 3.13 .venv`.
**Verified by:** `.venv/bin/python -m pytest sb/tests -q` → 62 passed.
**Also done:** `run.sh` now names an explicit 3.13 in its setup hint and refuses to run a
venv below 3.10 with a message that says to delete and rebuild (a venv cannot be upgraded
in place). `BENCHMARK_RESULTS.md` §4's setup line corrected.

### E-2 · `mcp` 2.0 broke the tool server
**Status:** verified.
`requirements.txt` pinned `mcp>=1.0.0` with no upper bound. The MCP Python SDK shipped
2.0.0, which renamed `FastMCP` to `MCPServer` and moved it from `mcp.server.fastmcp` to
`mcp.server.mcpserver`. A fresh install therefore fails at
`sb/live/mcp_app.py:25` with `ModuleNotFoundError: No module named 'mcp.server.fastmcp'`,
the MCP server never binds, and the runner aborts before reaching any model.
**Fix:** pin `mcp>=1.0.0,<2.0.0` (resolves to 1.29.0).
**Verified by:** `from mcp.server.fastmcp import FastMCP` imports; live smoke reaches the
model and the model makes real MCP tool calls.
**Deliberately not done:** porting `mcp_app.py` to the 2.0 `MCPServer` API. That would
change the tool surface the model under test sees, which is a benchmark change and needs
its own verification, not a dependency bump.

### E-3 · MCP startup failures were unreadable
**Status:** verified.
`start_mcp` passed `stderr=subprocess.DEVNULL`, so E-2's traceback surfaced only as
`RuntimeError: mcp server did not come up`.
**Fix:** keep stderr, poll `proc.poll()`, and put the child's traceback in the raised
error.
**Verified by:** this is how E-2 was diagnosed.

### M-1 · `run.sh` silently discarded every flag
**Status:** verified. **This is the reported bug.**
`run.sh:26` @`24331fb` read `live) exec "$PY" -m sb.live.runner ;;` — no `shift`, no `"$@"`. The
`demo` and `test` branches both forward arguments; `live` did not. So
`./run.sh live --model claude-opus-5 --seed 7` executed the runner with **zero
arguments** and fell back to argparse defaults at `runner.py:499-514` @`24331fb`:
`--model claude-haiku-4-5`, `--seed 42`, `--days 60`, `--daily-max 5`,
`--corpus corpus`. That is exactly the report: an unknown default model regardless of
the flag, and an identical configuration — hence an identical score — every run.
The trap only fires on `./run.sh live --model X`; `./run.sh --model X` falls through to
the `*)` branch and always worked, which is why it went unnoticed.
**Fix:** `live) shift || true; exec "$PY" -m sb.live.runner "$@" ;;`. The `|| true` is
load-bearing: a bare `./run.sh` has nothing to shift and `set -e` would abort.
**Verified by:** `./run.sh live --help` now reaches argparse; two live smokes launched
via `./run.sh live --model X` served two different models (see M-3).

### M-2 · the runner never reported the model actually served
**Status:** verified.
The header printed the model that was *requested*. The `system`/`init` event in the
CLI's stream already carries the model that was *loaded*, and `_parse_stream` read that
event but kept only `session_id`. A silent substitution was therefore unobservable, in
the live output and in every saved log in `outputs/` and `past/`.
**Fix:** `_parse_stream` and `_parse_codex` now return a 5th value, the resolved model.
The runner prints it once, flags a mismatch in red, warns on mid-run drift (M-3), and
stamps the served model plus seed, levers, and corpus next to the final score so a saved
log cannot be mislabelled.
**Verified by:** parser unit-checked on synthetic streams for both drivers; live smokes
print `model served claude-opus-4-8 ✓` and `claude-sonnet-4-5-20250929 ✓`.
Note the match test is `startswith`, because the CLI resolves an alias to a dated
snapshot (`claude-sonnet-4-5` → `claude-sonnet-4-5-20250929`).
> **Phase A amendment (C-7):** the stamp is weaker than this reads. `runner.py:530` falls
> back to the *requested* model when nothing resolved, prints it under the word "served",
> and a mid-run drift overwrites the value so the footer names only the last model.

### M-3 · `--model` across `--resume`
**Status:** verified.
A one-turn smoke cannot see this: if `--model` were ignored on resume, day 1 would be
correct and every later day would drift to the account default, producing exactly the
"same result every time" symptom even with M-1 fixed.
**Fix:** the resolved model is checked on **every** turn, not just the first, and any
change prints a `MODEL DRIFT` warning naming the day.
**Verified by:** a 3-day / 9-email smoke (`--limit 9`, two `--resume` turns) served
`claude-opus-4-8` on every turn with no drift warning. `--model` does survive `--resume`
on CLI 2.1.233, so M-1 was the whole story.

### M-4 · roster staleness
**Status:** closed, not a problem. `claude-opus-4-8`, `claude-sonnet-4-5` and
`claude-haiku-4-5` are all current, active model IDs. The roster in
`BENCHMARK_RESULTS.md` §2 resolves fine; the OpenAI half is untestable here only because
`codex` is not installed on this machine.

### M-5 · inherited environment
**Status:** open, unquantified.
`runner.py` spawns the CLI with the parent environment. A shell inside a Claude Code
session carries `CLAUDECODE`, `CLAUDE_CODE_SESSION_ID`, `CLAUDE_CODE_ENTRYPOINT` and
`CLAUDE_CODE_CHILD_SESSION`. Also `~/.claude/settings.json` sets `"model": "opus[1m]"`
globally, so any path where `--model` fails to reach the CLI inherits Opus rather than
erroring. Not observed to cause a problem — the smokes served the requested models — but
runs launched from inside an agent session are not obviously equivalent to runs launched
from a plain shell. Worth settling before the phase 7 comparison run.
> **Phase A note (O-8 opt 3):** the cheapest settlement is environment provenance stamped
> once per run (`claude --version`, `git rev-parse HEAD`, UTC start/end, hostname). That
> is O work and lands in phase 6.

---

## Corpus health check (phase 0 side-effect)

`.venv/bin/python -m sb.scale --filler 0 --seed 42 --days 200 --dst build/scaled0`

```
scaled corpus: 167 emails, 167 served over 57 days
filler: 0 junk emails burying 24 authored needle(s)
needle span: max 83, mean 31.6  (n=24)
oracle: 167/167 = 100%
```

The corpus lints clean and every answer key is satisfiable. The 57-day span matches the
`outputs/opus.md` and `outputs/sonnet.md` logs exactly, confirming those ran at default
levers (`daily_max=5`) rather than the `daily_max=21` pinned in `BENCHMARK_RESULTS.md`.
Mean needle span of 31.6 is low, which is the V-category concern: the retrieval axis the
benchmark is built to measure is barely exercised by the authored corpus alone.

> **Phase A corrections to the paragraph above. The measurement is reproduced; three of
> its readings are not safe.**
> - **"every answer key is satisfiable" holds only at the levers tested.** `sb/scale.py:111`
>   calls `build_plan` with no `levers=` argument and `sb.scale` exposes no lever flags, so
>   the gate always runs at `Levers(1,5,7)`. Across a sweep of 93 feasible settings, **12
>   score 166/167** (C-3), always failing
>   `Marketing-campaign-new-product-delay.serena-williams-reschedule`.
> - **"the corpus lints clean" is a statement about shape, not about gradeability** (G-4,
>   G-5, K's framing). The oracle titles every object with the answer key's own keywords
>   (`sb/oracle.py:52`), so it passes any corpus however unmatchable its keywords are.
> - **The lever inference is now an identification, not an inference** (C-1): replaying the
>   167 printed `· served <date>` pairs against rebuilt plans matches exactly one of 785
>   feasible lever combinations for opus and sonnet — `daily_min=1 daily_max=5
>   urgency_horizon=7`. For haiku seven combinations fit (`urgency_horizon ∈ 1..7`).
> - **The 31.6 mean is not "low", it is disqualifying** (V-1): the whole corpus is ~10.9k
>   tokens and the largest needle gap ~5.4k tokens of mail, against a ≥200k context window.

---

## Phase A — the 50 findings

Format per finding: **ID · title** — severity · cost to verify · status. Then the problem
and its single best sourced fact, then the options in one line. **The full evidence,
every command, and the open questions are in `docs/_repair/<letter>.md`.**

### G — grader identity (`docs/_repair/G.md`)

**G-1 · the match keyword is often not derivable from any text the model reads** —
blocks-measurement · free-offline · open
`_title_hit` (`sb/grader.py:68-70`) requires every keyword in `op.match` to appear as a
lowercased substring of `f"{obj.title} {obj.description}"`. For **32 of 125** graded
create/move ops (26%) the keyword appears nowhere in the rendered mail of the entire node.
Pooled over three runs at op level: keyword present in the node's mail → **146/303 = 48%**
pass; keyword absent → **8/99 = 8%**. The authored keyword predicts the outcome far more
strongly than the model does (43% / 50% / 51% across opus / sonnet / haiku). A second class
is created by rendering: `{!name = expr}` discards the anchor *name* (`sb/resolver.py:345-356`),
so `delayed`, `signoff`, `conference` exist in the JSON and in zero rendered bodies. The
The model is never told the matching rule. `SYSTEM_PROMPT` (`sb/live/runner.py:87-107`) is
silent on titles and on `description`; it *does* state the one-object rule in plain English
(`:107` ends "Never leave duplicates"). That makes the finding sharper, not weaker: the model
was told the right rule and is penalised by a different one, since G-2 shows the count check
fires on *distinct* obligations sharing vocabulary in 57 of 57 cases. What the prompt never
says is that the grader decides sameness by keyword substring over `title + description`.
*Options:* publish the contract to the model (changes the tool surface) · lint keywords
against rendered mail (32-op corpus edit) · semantic / LLM-judge matching (costs money,
nondeterministic) · grade by attribution instead of title (moves the problem into A).

**G-2 · the exactly-one rule counts keyword hits, not obligations** —
blocks-measurement · free-offline · open
`sb/grader.py:164-165` sets `count_ok = len(title_set) == 1` over the node's cumulative
same-kind pool. Across all four recorded runs, **not one** of the 57
`found N matching, expected exactly 1 (duplicate / double-booked)` failures involves two
objects with the same title (opus 0/14, sonnet 0/12, haiku 0/17, sonnet176 0/14). Every one
is a collision between *distinct* obligations sharing storyline vocabulary (`atlas`,
`reveal`, `launch`, `design`, `sponsor`). The log's wording is the opposite of what
happened, and readers will conclude models leave duplicates: they did not, in 57/57 cases.
Under a rule of "at least one matching object on an expected day" — restricted to `eq`/`any_of`
predicates where the day is unambiguous, a restriction the numbers move 1-2 points without —
opus 90→97, sonnet
91→98, haiku 98→109, sonnet176 102→111.
*Options:* best-match assignment (Hungarian/greedy; real algorithm change) · scope
exactly-one to the turn · relax to "≥1 satisfying, no stale survivor" (needs object
identity the design avoids) · leave the rule and fix corpus vocabulary (pushes authors
toward the artificial keywords that produced G-1).

**G-3 · `description` is in the match haystack, invisible in the log, unmentioned to the
model** — blocks-measurement · free-offline · open
`sb/grader.py:69` builds the haystack as `f"{obj.title} {obj.description}".lower()`;
`_fmt_obj` (`sb/grader.py:125-129`) prints only the title; `SYSTEM_PROMPT` never mentions
the field, which `mcp_app.py:42,73` exposes. In `outputs/opus.md`, **23 of 101** attributed
objects (23%) have a title lacking the keyword and matched through the description alone
(sonnet 14/97, haiku 8/97, sonnet176 2/91). *An earlier draft added "the spread tracks how
verbose each model is, so the metric is partly measuring writing style". That clause is
dropped: descriptions are never rendered (`grader.py:125-129`, and O-5 says the same), so
description verbosity cannot be measured from any surviving artifact, and the one verbosity
proxy that does exist runs the other way — mean title length is opus 47 > haiku 40 > sonnet
36 while description-only matches are opus 23 > sonnet 14 > haiku 8. O-3's state dump would
settle it.* Controlled offline experiment: take the
**unmodified oracle** and add a realistic description (subject + first 180 chars of the
rendered body) to every object; score falls **167/167 → 141/167 (84%)**, 25 of the 29 lost
details being collisions, with nothing about the scheduling changed.
*Options:* match on title only · require the keyword in the title, description breaks ties
only · keep it and tell the model (creates a keyword-stuffing gaming surface) · print the
description in `_fmt_obj` (diagnostic only, cannot move a score).

**G-4 · the oracle certifies satisfiability and is structurally blind to gradeability** —
blocks-measurement · free-offline · open
`sb/oracle.py:52` titles every object `" ".join(op.match)`, so the reference model writes
the answer key's own keywords into the calendar and scores 100% on any corpus, however
unmatchable. `sb/scale.py:126-127` prints that 100% as "must be 100% — corpus is valid at
scale". Title-policy sweep through the *shipped* `sb.engine.run` + `sb.grader`, with the
agent scheduling-perfect in every variant:

| title policy | score | action-emails only (n=111) |
|---|---|---|
| P0 `" ".join(op.match)` — shipped oracle | 167/167 **100%** | 111/111 100% |
| P1 `op.name` verbatim | 160/167 96% | 104/111 94% |
| P2 `op.name` humanized | 157/167 94% | 101/111 91% |
| P3 the email's subject line | 92/167 **55%** | 36/111 32% |
| P4 subject + humanized name | 140/167 84% | 84/111 76% |
| P0 + a realistic description | 141/167 84% | 85/111 77% |

Recorded models: opus 54%, sonnet 54%, haiku 59%. **A scheduling-perfect agent that titles
calendar entries after the email subject scores inside the band of all three real runs.**
Even the author's own name for the obligation costs 6–10 points.
*Options:* add a paraphrase-oracle gate at P2/P3 with a CI threshold · replay the ~270
titles real models actually produced through the grader (free; biased toward objects that
already matched) · randomize the title within the contract (catches nothing about
discoverability) · leave the oracle and rename its output line.

**G-5 · lint check #5 compares author strings to author strings** —
distorts-measurement · free-offline · open
`sb/schema.py:551-572` flags a collision only when one obligation's `match` keywords are
substrings of a sibling's `match` keywords. Both sides are author-invented; the grader
matches titles the *model* invents — which the check's own comment at `:556-558`
acknowledges and then does not act on. Shipped check: 0 flags. A variant comparing each
`match` set against its siblings' obligation **names**: **10 flags in 7 of 15 nodes**
(`['atlas']` alone catches three of its four sibling names), and every one appears as a
real `found N matching` failure in the logs. Separately, 23 keywords are substrings of
ordinary words in their own node's mail (`'ai'` in `email`/`openai`, `'end'` in `attend`,
`'sign'` in `design`). The check is restricted to `verb == "create"` (`:563`), so cancel
keywords — the verb most sensitive to collisions (G-7) — are never examined.
*Options:* widen the probe set to names + subject words · lint against a checked-in file of
observed model titles · downgrade to a counted warning · delete #5 and rely on whatever
replaces the exactly-one rule.

**G-6 · the pool is cumulative over the whole run, so obligations get harder by position** —
distorts-measurement · free-offline · open
`_grade_op` matches against the entire node pool (`sb/grader.py:151`), which
`sb/live/runner.py:137-146` builds from the whole store filtered only by node. Pass rate by
pool depth — bucketing by **prior same-kind obligations in the answer key**, not by prior objects
the model created, which is the reading that reproduces — pooled over three runs: 0 prior **42/81 = 52%**; 1–2
**65/156 = 42%**; 3–5 **33/111 = 30%**; 6+ **14/54 = 26%**. `_node_state`'s fourth parameter
`sid_filter` is accepted and never referenced (`grep -rn "sid_filter" sb/` returns only the
signature) while the call site at `runner.py:509` passes `eid_new`, so the code reads as if
the pool were turn-scoped when it is not.
*Options:* keep cumulative and report the decay as a caveat on every score · window the pool
by N days · scope to the obligation's lifetime (needs bookkeeping the design avoids) · fix
the count rule instead (G-2) and assume pool depth stops mattering.

**G-7 · `cancel` is graded as keyword absence over the cumulative pool** —
distorts-measurement · free-offline · open
`sb/grader.py:155-160` passes a `cancel` only when **zero** objects in the node's cumulative
pool contain the keyword, so any sibling the model was supposed to *keep* makes the cancel
unpassable. All 13 cancel failures across the four runs are of this form: `~"launch"`
survived as `"Project Atlas — public launch"`; `~"design"` as `"Design team future
discussion with Melissa"`; `~"dynamics"` as `"Boston tech trip (WHOOP + Boston Dynamics)"`
in **all four runs** — the models modelled one trip that the corpus splits into five
obligations. Pass rates 6/9, 5/9, 5/9, 6/9. The oracle cannot reproduce this
(`sb/oracle.py:54-56` deletes by the string it wrote).
*Options:* grade cancel through the same assignment as create/move (depends on G-2) ·
require absence only among objects previously attributed to that obligation · scope absence
to the expected date · treat the Boston case as a K-category granularity problem.

**G-8 · `match` defaults to the whole obligation name as one contiguous phrase** —
distorts-measurement · free-offline · open
`sb/schema.py:154-156` sets `match=list(raw.get("match", [])) or [name]`. **21 of 134 ops
(16%) take the default**, and they pass at **8/63 = 13%** against **146/339 = 43%** for
explicit keywords. By keyword length: ≤5 chars 35%, 6–9 chars 45%, 10–15 chars 41%,
**≥16 chars (phrases) 3/42 = 7%**. All 134 ops have exactly one keyword, so the
multi-keyword conjunction the schema offers for disambiguation is entirely unused.
`_parse_op` (`sb/schema.py:130-156`) performs **no** validation of `match`: no non-empty
check, no whitespace strip — `Giano Ronaldo marketing campaign ` carries a trailing space
and works only because the haystack always appends one.
*Options:* make `match` required (schema change, webapp editor consequence) · tokenise the
default name into a keyword list · lint the field (strip, reject empty, warn on length) ·
treat the 21 ops as a K-category corpus edit so a later score move stays attributable.

**G-9 · the grader's own explanation cannot distinguish its failure modes** —
slows-work · free-offline · open
`sb/grader.py:168` and `:172-173` fire under the *same* guard (`not title_set`), so the
brief's "all 52 opus cases read `(nothing matching created)`" is a tautology of the code,
not an observation. `_fmt_obj` prints neither the description that decided the match, nor
the kind, nor the `email_id`. Six causes are reachable and none is separable from a log:
under-action; title/description lacking the keyword; kind mismatch (`grader.py:151`);
`email_id` naming another node; `email_id` resolving to nothing; action on a later day
(grading is one-shot at `runner.py:503-511`).
*Options:* emit a machine-readable per-op record (this is O's work) · split the reason
strings by cause inside `_grade_op` · print description + `email_id` in `_fmt_obj` (changes
the format `analyze.py:25` re-parses) · do nothing and rely on phase 6's state dump.

**G-10 · the identity logic has no unit tests** — slows-work · free-offline · open
`sb/tests/test_grader.py` contains four tests, all about date predicates and oracle blackout
handling. Nothing tests `_title_hit`, the description haystack, the kind filter, the
exactly-one rule or the cancel rule — the code responsible for ~85% of recorded failures.
Worse, `passed = count_ok and len(matched) >= 1` (`grader.py:165`) reaches the date
predicate only when `len(title_set) == 1`, so **on 47–55% of graded obligations the answer
key's date is never evaluated at all**: the outcome is settled by string matching before any
temporal reasoning is tested. `op.tolerance` is `exact_day` on 134/134 ops.
*Options:* characterisation tests pinning today's behaviour first · property tests stating
the intended contract (requires deciding the contract, which is G-2) · golden-file
regression on a fixture corpus + store.

### A — attribution and no-action (`docs/_repair/A.md`)

Framing that all six A findings rest on: the model's `email_id` stamp reaches the score
through exactly two channels. `runner.py:141-143` keeps an object only if
`corpus.emails.get(obj.email_id).node == node`, so the stamp selects the node pool and an
unknown id drops the object from every pool. `runner.py:507-510` builds `eid_new` and
passes it as `TurnDelta`, which `grader.py:186-195` reads **only** in the no-action branch.
For the 111 emails carrying ops the delta is never read; for the 56 no-action emails it is
the whole grade.

**A-1 · more than half of every score is a no-action verdict, and the ranking is not robust
to the attribution rule** — blocks-measurement · free-offline · open
No-action passes are 51/90, 51/91, 56/98 of each run's passes — **56–57% of every headline
score is the abstain check**, whose only input is the model's self-reported `email_id` plus
same-day timing. Re-grading the same logs day-scoped (any object created that day counts
against every no-action email in the batch) gives opus 90→55, sonnet 91→57, haiku 98→43.
56 of 167 emails carry no ops and **31 of those 56 are in `Innovation-comp`**.
*Options:* report no-action and acting accuracy separately (the headline number disappears)
· reweight the two classes explicitly · rebalance the corpus · fix only the attribution rule
(A-3) and re-measure.
> **See "Contradiction 1" below.** A-1's rank inversion is confounded with K-2's lever
> difference and must not be cited on its own as evidence about attribution.

**A-2 · the live attribution split is untested and semantically different from the offline
path the oracle validates** — blocks-measurement · free-offline · open
Attribution is implemented twice. Offline, `engine.py:143-147` grades one email per turn
with an unfiltered delta and never uses `email_id` to route. Live, `runner.py:501-510`
grades one *day* per turn and splits by the model-supplied stamp. Same input, divergent
failure: `engine.py:97` raises `KeyError` on an unknown id; `runner.py:141-143` returns
`None` and silently drops the object. `grep -rn "_node_state\|_turn_delta\|by_eid\|eid_new"
sb/` returns only `sb/live/runner.py` — **no test touches the live path**, and
`sb/tests/test_e2e.py:9` imports the offline engine. `oracle: 167/167 = 100%` is therefore
evidence about a code path no live run uses.
*Options:* delete `sid_filter` and unit-test the live helpers (risks entrenching the rule
A-3 questions) · collapse the two paths onto one implementation · make the offline path
day-based so the oracle exercises attribution (still would not exercise a *wrong* stamp) ·
document the divergence and stop citing oracle 100% as evidence about the harness.

**A-3 · the no-action check reads a day-scoped and id-scoped delta, giving five escapes** —
distorts-measurement · free-offline · open
`grade_email` fails a no-action email only if `turn` is non-empty, where `turn` is objects
created in *this day's single turn* stamped with *this email's exact id*. Five routes out:
(1) sibling stamp same day (`runner.py:507`) — 46 of the 56 no-action emails share a day
with an acting email; (2) **action on a later day** (`runner.py:501` vs `:431`, verdict
written once at `:503-512`) — this is the route actually observed, see A-6; (3) unknown or
stale stamp (`store_app.py:89-92` warns and does not block); (4) create-then-delete inside
one turn, invisible to a set difference over end-of-turn state; (5) **`patch_event`
(`store_app.py:139-144`) mutates in place, keeps the id, and calls no attribution check** —
a model can retitle and redate a stale object with zero grading consequence. Counter-evidence
stated plainly: all 10 observed over-action failures carry titles topically derived from the
no-action email itself, so where models over-acted they stamped correctly; there is no log
evidence of an adversarial stamp in any run.
*Options:* re-evaluate no-action at end of run (closes 2 and 3) · grade the day's objects as
one pool against the day's batch (closes 1 and 2, makes the metric day-granular, moves 34–55
verdicts per run) · grade from an action log (closes 4 and 5, partly reverses the state-based
design) · reject `create_*` whose `email_id` is not in today's batch (changes the tool
contract; would have changed exactly one verdict across all three runs).

**A-4 · `_watch_attribution` is blind to the exact case it is cited as evidence about** —
distorts-measurement · free-offline · open
`store_app.py:86-92` warns on `invalid_email_id` (never served) or `stale_email_id` (served
earlier). **There is no third branch**: an id belonging to a different email in *today's*
batch satisfies neither condition and produces no warning. That is precisely the sibling
stamp the brief's §2.6 cites the warning count as evidence about. The store is also
asymmetric — `store_app.py:225-230` 404s on an unknown id at *read*, while `:126-131` and
`:156-161` accept any string at *write*. Nothing aggregates: `runner.py:489-490` prints the
day's warnings, the store is terminated with no dump, `sb/analyze.py` never reads a warning
line, and the end-of-run summary has no warning count. Transcription hazard the monitor
would catch and nothing measures: email ids average 45 chars (max 77), 108 of 167 exceed 40
chars, and there are 16 prefix *pairs* (one id a strict prefix of another) spanning 14 distinct ids.
*Options:* add a `sibling_email_id` warning (may only be expressible as a node-mismatch
heuristic) · warn when a stamp resolves to a node with no op the object could serve (leaks
answer-key structure into the store) · persist warnings into the artifact and count them ·
validate `email_id` on write as it already is on read (changes the tool contract mid-benchmark).

**A-5 · a wrong `email_id` silently deletes the object from every pool** —
distorts-measurement · **needs-one-live-run** · open
`runner.py:137-146` returns `o if (e and e.node == node) else None`. An unknown id makes the
object invisible in `NodeState` for every node *and* excluded from every `TurnDelta`, while
remaining visible to the model through `list_events`. The honest model that does the right
work and mis-stamps gets `(nothing matching created)` and looks lazy. Exposure: 150 of 167
emails have a same-day sibling in a *different* node; 51 of opus's 52 and 46 of sonnet's 47
"(nothing matching created)" details fall on multi-email days, exactly where a wrong stamp is
silent. **The evidence here is thin and A says so:** there is no observed instance of
cross-node misattribution in any run — every object in every duplicate-failure `actual` field
is topically native to its node — but the log also cannot refute it, because an object
misattributed into a node where it matches no keyword never appears in any `actual` field.
The case rests on the A-4 blind spot plus exposure counts, not on observation.
*Options:* grade against all objects regardless of stamp (makes cross-node collisions worse)
· fall back to a day-scoped pool for unresolvable stamps (perverse incentive) · reject
unresolvable stamps at the store · instrument and measure it (the only option that produces
the missing number; needs O).

**A-6 · worked case: one stale stamp produced one false PASS and cost one earned PASS** —
distorts-measurement · free-offline · open
The single attribution warning in the entire recorded corpus of runs is not benign.
`outputs/opus.md:16-19`: `Innovation-comp.sponsor-mixer-before-the-final` serves on day 1
with no ops and passes "correctly took no action". `outputs/opus.md:67`, day 4: the model
creates `"Sponsor mixer with retail partners (optional appearance)"` stamped with that day-1
id (day 4's batch is a single email, so this is route 2, not a sibling stamp). The day-1
verdict is never revisited. `outputs/opus.md:689-692`, day 35: the mixer is the only object
surviving the `~"sponsor"` cancel, failing an unrelated email. **The two errors run opposite
and net to zero on the headline**, which is why the anomaly is invisible in the aggregate
while both verdicts are wrong. Counterfactuals: object never existed → 90→91; no-action
re-evaluated at end of run → 90→89. This contradicts the brief's reading of the same warning
as "a theoretical hole, not an observed exploit": it is observed and score-changing, though
not an exploit — the model was acting honestly.
*Options:* treat as the `POTENTIAL_GAMING.md` hardening trigger (n=1 is not "regularly") ·
treat as a corpus problem (the "optional appearance" key may be wrong) · use it as the seed
case for the hand-grade · wait for O.

### O — run artifacts and observability (`docs/_repair/O.md`)

**O-1 · the harness writes nothing; the surviving artifacts are hand-copied stdout** —
blocks-measurement · free-offline · open
`sb/live/runner.py:535-558` is `main()`'s full argparse surface and has no output flag; there
is no `open(..., "w")` in the module. `sb/live/runner.py:513-516` terminates the store with
nothing read out of it, and `sb/live/store_app.py:22-27` shows the entire run state is four
module-level in-memory objects in a uvicorn child. Capture is shell redirection into
`build/`, which `.gitignore` excludes. `past/claude-haiku-4-5.md:7` cites a run log that does
not exist; `past/claude-opus-4-8.md` is 0 bytes. **Consequence, in its defensible form: no
run already paid for can be re-graded under a rule that would newly admit objects the log
never rendered** — anything widening `_title_hit`, changing `op.kind`, or re-attributing by
`email_id` — because `grader.py:151-152` filtered those out before printing. G-1, G-3 and G-8
are blocked by this. Rules that relax the count check, re-scope the no-action delta, or drop
an object *are* evaluable offline today, and this register does exactly that three times
(G-2, A-1, A-6). An earlier draft of this finding said "can ever be re-graded" flatly; the
verifier reproduced two of the three counter-examples from the logs alone.
*Options:* `--out DIR` writing a JSON/JSONL tree · persist the store instead of the runner ·
keep stdout as the only artifact but make it machine-parseable · instrument future runs only.

**C-1b · the model label on every surviving artifact is an unverified assertion** —
blocks-measurement · free-offline · open · **phase 1** *(added 2026-08-18)*
The four artifacts' headers name a model (`outputs/opus.md:3` reads
`model claude-opus-4-8 via claude`). Per M-2 that is the model **requested**, not served;
the runner did not report the served model until phase 0, on 2026-08-17. The recorded runs
are dated 2026-07-04 and 2026-07-26, so **M-1 was live when they ran** — `./run.sh live
--model X` discarded every flag and fell back to `claude-haiku-4-5`. If they were launched
that way, the labels are wrong.

*Measured this turn, and it partly rescues them.* The broken path also forces the argparse
defaults, including `daily_max=5`. `past/claude-haiku-4-5.md` ran `daily_max=21`, so it
**cannot** have gone through the trap and its flags demonstrably worked. `outputs/opus.md`
and `outputs/sonnet.md` both ran `daily_max=5`, which is what the trap produces, so both
are consistent with it — but if *both* had been trapped they would be one model at one
config, and they are not: 10/94 byte-identical titles (10.6%), em-dash rate 62.5% vs 0.0%,
mean title length 45.0 vs 35.9. So **at most one of the two can be mislabelled.** Testing
the candidate: sonnet-vs-haiku is 14/87 identical (16.1%) with mean title length 35.9 vs
40.3 — elevated over the 10.6% baseline but not a same-model signature.

Two things *are* verified, by independent checks: every one of the 167 verdicts in each of
the three current logs resolves to an email id that exists exactly in today's corpus (0
orphans; `past/claude-sonnet-4-5.md` has 9 orphans from the retired `Company_Retreat` node,
correctly identifying it as the older corpus), and C-1's serve-plan replay is unique. The
corpus binding is sound. The model labels are not.
**Verdict: probably right, not provable.** No artifact can settle it — the served-model
event was discarded pre-phase-0, `build/*.log` is gone (verified absent on disk), and the
header corpus-sha algorithm exists nowhere in the repo (C-5). Only a fresh run can, which
is what phase 1c is for.
*Options:* record the labels as asserted-not-verified and move on · re-run one model at the
historical config and compare (this is 1c) · drop the historical runs from any published
comparison · attempt deeper stylometry (weak; cannot certify a specific model without a
known-model reference sample).

**O-2 · the tool trace drops calls by message-id collapse, and the loss is model-dependent** —
blocks-measurement · free-offline · open
`sb/live/runner.py:183` — `seen[msg.get("id", id(ev))] = msg` — is last-write-wins per
message id, so a parallel tool-call batch loses every block but the last. Confirmed against
the real parser on synthetic streams: five `get_email` blocks in five events under one id
return one call; a `tool_use` followed by `text` under one id returns **zero** tools. A second
independent bug: the fallback key `id(ev)` is the CPython address of a dict rebound each
iteration, so the allocator reuses it — six id-less messages return **4**, deterministically,
and six with `"id": null` return **1**. The `codex` driver does not dedupe
(`runner.py:355-357`), so trace fidelity differs by driver by construction.
Sonnet's per-day `get_email` histogram is `{1: 57}` — exactly one on every one of 57 days,
with 1 to 5 emails arriving — against opus's `{1:11, 2:13, 3:12, 4:12, 5:9}`.
`sb/analyze.py:51-58` derives its entire `searched` signal from this trace.
*Options:* persist raw CLI stdout per turn and parse offline · fix the accumulator to append
and drop the `id(ev)` fallback · move the record server-side (blind to CLI built-ins) · do
both client and server and cross-check, which measures the loss directly.
> **See "Contradiction 2".** O's claim that opus lost `create_*` on 18/57 days does not
> survive; the rest of O-2 does.

**O-3 · a final-state dump is not sufficient: the store records no history and no day** —
blocks-measurement · free-offline · open
Grading is incremental per simulated day against the node state *as it stood that day*
(`runner.py:431`, `:500-501`, `:507-510`). Store objects carry no creation day
(`store_app.py:32-51`, `:130`, `:160`), updates overwrite in place with no history (`:143`,
`:173`), and deletes are destructive (`:149`, `:180`) — so a `cancel`, which grades on
absence, leaves no trace at all. A title that matched on day 10 and was renamed on day 40
grades differently on day 10 than a final dump would say. What need *not* be captured
object-by-object: the serve plan and every rendered body are deterministic from
`(corpus, start, seed, levers)` (`sb/scheduler.py:15`). Note also that no corpus-hash code
exists anywhere in the repo, yet two `past/` headers record one (see C-5).
*Options:* runner-side per-day snapshot of `/state` (misses intra-day ordering and deletions)
· store-side append-only mutation log with a day stamp (strictly richer; store schema change)
· add `created_on`/`updated_on` and dump once (still loses deletions and pre-rename titles) ·
embed the corpus itself rather than a hash.

**O-4 · retrieval is unobservable: the store logs no reads** —
blocks-measurement · free-offline · open
`store_app.py:192-213` (`/inbox/search`) writes only two coarse warnings — `broad_search`
when neither `q` nor `sender` is given, `large_search_result` when `limit == 0 or > 25`. A
normal targeted search, the behaviour the benchmark exists to measure, is recorded nowhere.
`:216-222` and `:225-230` are pure reads with no counter. By contrast every write path calls
`_watch_attribution`. `mcp_app.py:112-134` is a thin pass-through with no logging either.
**This is the one item in O that a state dump cannot fix**: retrieval leaves no server-side
residue, so instrumenting it means changing a model-facing surface — the same objection that
blocked the mcp 2.0 port (E-2). That decision cannot be deferred to "just dump the state",
and it is why O sits immediately before the paid run rather than after it.
*Options:* log every store read with day and params · log at the MCP layer instead (separate
process, needs its own artifact path) · count only, no arguments (fixes the `searched` boolean,
loses *what* was searched) · rely on a fixed client-side parse and change nothing model-facing.

**O-5 · the log renders only keyword-matched objects, so the dominant failure is
unfalsifiable** — blocks-measurement · free-offline · open
`sb/grader.py:151-152` filters the pool by kind *and* keyword before anything is rendered;
`:168` prints `(nothing matching created)` when that set is empty; `:190`'s no-action branch
renders only objects stamped with this email's id. `runner.py:80-85` prints nothing else about
the store. Distinct object titles ever visible in an `actual` field: 71 / 70 / 68 / 65 across
the four logs — everything above that is invisible in every artifact. **The correct statement
of brief §4.3 is therefore not "unmeasured" but "unmeasurable from any surviving artifact".**
*Options:* render the unmatched pool in the machine artifact · emit per-op diagnostic codes
(pre-commits to a taxonomy that is a phase-2 decision) · dump raw state and compute
diagnostics offline (keeps the grader untouched) · accept §4.3 stays open until phase 7.

**O-6 · tools and narration are attributed per day, never per email, and narration is
truncated** — distorts-measurement · free-offline · open
`runner.py:57-66` prints one tools line and one narration snippet per day, and `:65` cuts the
narration at 200 characters. Measured: snippets sitting at the cap on **57/57 opus days, 50/57
sonnet, 16/16 haiku, 18/19 past-sonnet**. Day-level attribution inflates the search signal:
opus's single day-1 `search_inbox` tags 5 of 167 emails as `searched=True`; at
`daily_max=21` one search on a 21-email day would tag 21. Latent parser fragility:
`sb/analyze.py:51` matches `if "tools" in line` before the PASS/FAIL check, so any line
containing the substring resets `day_searched` — not observed in any of the four logs.
*Options:* record tool calls with day index and arguments in the artifact · keep full narration
in the artifact and truncate only in the terminal · attribute by inspecting tool arguments
(works for everything except `search_inbox`, the one that matters) · retire `analyze.py`'s
log-scraping (breaks its ability to read the four historical logs).

**O-7 · infrastructure errors and retries fold into the score with no marker** —
distorts-measurement · free-offline · open
`runner.py:491-496` prints `ERROR after retries` on a line carrying no `PASS`/`FAIL` token, so
`analyze.py:25`'s regex never matches it and the email vanishes from analysis rather than being
counted. `runner.py:518-528` renders an errored email with the same red dot as a failed one.
Worse, `before = _all_ids(...)` is captured at `:431`, **outside** the retry loop at `:443-486`,
and there is no rollback — so every failed attempt's objects land in the same day's delta and
manufacture exactly the `found 2 matching, expected exactly 1` failure §2.3 counts 14 times for
opus. Latent, not observed: all four logs contain zero `retry` / `session limit` / `ERROR after`
lines — though `past/claude-haiku-4-5.md:6` records that the project has already spoiled one
run this way.
*Options:* tri-state per-email status plus an attempt count, and report the score with and
without errored emails · snapshot and roll back the store before each attempt · refuse to score
a run with any errored day · recompute `before` inside the retry loop (strands the failed
attempt's objects in the cumulative pool).

**O-8 · no cost, timing, token or version capture, although the CLI hands them over** —
slows-work · free-offline · open
`runner.py:184-186` reads the CLI's `result` event for `session_id` and `is_error` and drops
the rest. `grep -rn "total_cost_usd\|duration_ms\|num_turns\|usage" sb/` returns two hits, both
inside the rate-limit regex. No timestamp, no CLI version, no repo sha anywhere. The register's
whole sequencing argument is economic and the project cannot state what a run costs; M-3 was
settled against "CLI 2.1.233", a version no artifact records.
*Options:* keep the whole `result` event per turn · extract a fixed small set · stamp
environment provenance once per run (also settles M-5) · time turns in the runner rather than
trusting the CLI.

**O-9 · objects with an unparseable date are dropped from grading without a trace** —
distorts-measurement · free-offline · open
`runner.py:128-134` returns `None` when neither `start` nor `due_date` parses as ISO, and both
`_node_state` and `_turn_delta` drop it with no branch and no logging. The store accepts any
string (`store_app.py:32-51`, `:126-131`, `:156-161`). Since `grader.py:188-189` computes the
no-action verdict as `passed = not created`, a dropped object **turns an over-action into a
pass**. No warning covers it. Size unknown and unmeasurable from any existing artifact.
*Options:* validate at the store's write path (changes the tool surface and scores) · warn only,
matching the existing `_warn` precedent (zero score impact, quantifies the term) · count the
drops in the runner and surface them in the artifact · wait until it is shown non-zero.

### C — config and provenance (`docs/_repair/C.md`)

Phase 0's M-2 stamp covers: header `runner.py:403-408` (requested model, driver, seed, day
count, email count, start, `daily_min`-`daily_max`, `urgency_horizon`, corpus **path**) and
footer `runner.py:529-531` (served model, seed, `daily_min`-`daily_max`, email count, corpus
**path**). Stamped nowhere: corpus **content** identity, `urgency_horizon` in the footer,
`--start`/`--days`/`--limit`, the git revision of `sb/`, the driver CLI version, the
`SYSTEM_PROMPT` version, a timestamp. Neither line is machine-readable and neither appears in
any of the four saved artifacts, all of which predate phase 0.

**C-1 · no artifact records its levers; two documented commands overwrite the corpus they are
recovered from** — blocks-measurement · free-offline · open · **phase 1**
No saved artifact records its levers. They are nonetheless recoverable *today*: each log prints
its own serve plan (`runner.py:79`), and replaying the 167 `(email_id, serve_date)` pairs
against plans rebuilt from the current corpus identifies the levers uniquely — `outputs/opus.md`
and `outputs/sonnet.md` match **exactly one of 785 feasible combinations**
(`daily_min=1 daily_max=5 urgency_horizon=7`, plan digest `832b0b44bce0`). The 785 is the
count of schedulable settings over the swept grid `daily_min ∈ 1..3`, `daily_max ∈ 4..30`,
`urgency_horizon ∈ 1..7` — the number is meaningless without that grid, so it travels with
it. `past/claude-haiku-4-5.md` matches **seven** settings, not three: `daily_max=21` with
`urgency_horizon ∈ 1..7`, digest `4dd378ec78f1`. (An earlier draft said three; that was an
artifact of dropping the grid in the merge. The verifier re-swept it.) The recovery is
reproducible from git at any date — `corpus/` is tracked, clean and unchanged since
`24331fb` — so it is not perishable, but it is easier to record now than to re-derive. **The destruction path is documented, not
hypothetical:** `BENCHMARK_RESULTS.md:146-147` names `scripts/recover_corpus.py` (which fetches
production live at `:25,32`) and `scripts/fix_match.py` (`:21`, rewrites `corpus/nodes/` in
place) as the reproduce path.
*Options:* record the recovered provenance in this register now as a one-off note · write a
machine-readable sidecar per artifact before touching the corpus · gate corpus writes while an
unrecorded artifact exists · accept the loss and declare the three runs superseded.

**C-2 · `urgency_horizon` is absent from the score-line stamp and moves 89% of serve dates** —
distorts-measurement · free-offline · open
The footer (`runner.py:529-531`) records `daily_min`-`daily_max` and not `urgency_horizon`; the
hand-written `past/` headers omit it too. Measured at `daily_min=1, daily_max=5`, seed 42:
moving `urgency_horizon` 7→14 changes the serve date of **148 of 167 emails** (max shift 34
days), and 5→7 changes 119. All of `{5, 7, 14}` still produce exactly 57 non-empty days, so the
header's day count cannot distinguish them. By contrast `--days` is inert once feasible.
Answer keys resolve against the serve date (`runner.py:506`), so two identically-stamped runs
can have been asked different questions.
*Options:* add `urgency_horizon` and `--start` to the footer · print a derived plan digest
(identifies corpus and levers at once; says two runs differ without saying how) · emit the full
argparse namespace as JSON (should land with O's artifact, not separately) · pin the levers in
code and delete the flags.

**C-3 · corpus satisfiability is lever-dependent, and the oracle gate is hardcoded** —
distorts-measurement · free-offline · open
`sb/scale.py:111` calls `build_plan` with no `levers=` argument and `sb.scale` exposes no lever
flags, so the gate always runs at `Levers(1,5,7)`. Sweeping 93 feasible settings, **12 score
166/167**, always failing `Marketing-campaign-new-product-delay.serena-williams-reschedule`
with `found 2 matching, expected exactly 1`. A third hardcode exists at
`scripts/fix_match.py:114-115` (`n_days=730, Levers(daily_max=21)`), matching neither the
argparse defaults nor `sb.scale`. **Corollary that matters for G:** the oracle is a *partial*
detector for between-obligation keyword collisions, not a blind one — it is blind to keywords a
real model would never produce, which is a narrower claim than brief §2.5 makes.
*Options:* add lever flags to `sb.scale` · make the gate sweep a lever grid and report the worst
case · pin the levers as constants · treat it as a single corpus defect and fix that keyword.

**C-4 · `sb.analyze` takes the levers by hand and silently reports a different span grid** —
distorts-measurement · free-offline · open
`sb/analyze.py:88-90` rebuilds the serve plan from its own CLI defaults (`--corpus build/scaled`,
`--days 300`, `--daily-max 5`) and never reads the levers out of the log it is analysing.
`parse_log` extracts only PASS/FAIL and `searched`. Same haiku log, two lever settings, no
warning either time: the default invocation bins 19/5/0 needles, the correct one
(`--daily-max 21 --days 30`) bins 13/10/1 — the `100+` bin appears only in the correct run, and
the span-independent overall figure hides the error. `BENCHMARK_RESULTS.md:159-160` already
flags this as a manual discipline ("MUST pass the same levers"), which makes it documented and
unchecked.
*Options:* parse the levers from the runner's stamp and refuse on disagreement (works only for
post-phase-0 logs, i.e. none of the four) · consume O's machine-readable record instead · make
the flags required · verify the rebuilt plan against the `· served <date>` lines every log
already prints (self-checking, works on the four old logs).

**C-5 · corpus identity is asserted by a hash whose algorithm exists nowhere in the repo** —
distorts-measurement · free-offline · open
Four places record a truncated corpus sha (`BENCHMARK_RESULTS.md:81,205`,
`past/claude-sonnet-4-5.md:4`, `past/claude-haiku-4-5.md:4`), and
`grep -rn "sha256\|hexdigest\|blake2" scripts/*.py sb/*.py sb/live/*.py` returns nothing. Six
plausible reconstructions over `corpus/nodes/` all fail to reproduce `d737d44e14dc7d20`. What the
runner stamps instead is a *path* (`runner.py:531`), so two runs against completely different
corpora at the same path stamp identically. The evidence is bounded honestly: a seventh
algorithm (e.g. hashing the webapp export rather than the node files) could still match.
*Options:* define and implement one canonical digest and mark the historical values unverifiable
· stamp `git rev-parse HEAD:corpus` instead (useless for generated corpora) · stamp both · stamp
the derived plan digest (misses every `match`-keyword edit, which is exactly what phase 5 does).

**C-6 · the whole evidentiary base entered in one commit titled "." that also rewrote the
harness** — distorts-measurement · free-offline · open
`24331fb` (2026-07-26, message `.`, 26 files, +9257/-25) added `BENCHMARK_RESULTS.md`, the entire
corpus, all five run artifacts, both corpus-mutating scripts **and** a 190-line behavioural
rewrite of `sb/live/runner.py` (the codex driver, `Levers`, `_limit_reset_wait`, a rewritten
`build_cmd`). The four logs were produced by working-tree code that was never committed while in
use. The 176-email corpus was never committed at all — `git log --all -- corpus/` goes
`4f24e2a` (blank slate) → `24331fb` — so `past/claude-sonnet-4-5.md`'s 102/176 is attributable to
no corpus in this repo. `CLAUDE.md` carries the never-both-in-one-commit rule and is itself
gitignored, so the rule was never visible to the author who broke it.
*Options:* pin `24331fb` as the best-available harness revision for the three 167-corpus logs and
mark the fourth un-attributable · retire the pre-phase-0 artifacts as baselines · add a
commit-time guard against touching `corpus/` and `sb/` together · move the operating rules into a
tracked `CONTRIBUTING.md`.

**C-7 · the stamp prints "served <model>" from the requested model when nothing was served** —
distorts-measurement · free-offline · open
`runner.py:530` prints `served {resolved_model or model}`. `resolved_model` is `None` at `:410`
and set only inside the success branch at `:456-458`, so an all-errored run — not hypothetical,
`BENCHMARK_RESULTS.md:173` records a run that lost 35 emails to a usage window — prints the
requested model as though observed. Drift collapses to the last value (`:465-469`), so the footer
names one model for a run that used two while the warning sits in the middle of a 1100-line log.
The mismatch test `rmodel == model or rmodel.startswith(model)` (`:459`) is prefix-loose in both
directions.
*Options:* print `served: not observed (requested X)` when unresolved · record and print the full
set with day ranges · make an unresolved model or any drift a hard abort (throws away a paid
partial run) · leave it and treat the footer as a summary.

**C-8 · run logs default to a gitignored path, so artifacts are preserved only by hand** —
slows-work · free-offline · open
Every documented run command redirects into `build/` (`BENCHMARK_RESULTS.md:152-160`,
`RUNNING.md:78`, `RUN_RESULTS.md:12`), which `.gitignore:9,39` excludes. `find . -name "*.log"`
returns nothing. Three known gaps are one mechanism: the 0-byte `past/claude-opus-4-8.md` (born
empty in `24331fb`), the 84/176 haiku run **overwritten** by a later run of the same model at the
same model-id filename, and the ~51% run of brief §4.4. Two artifact conventions coexist in one
commit, and the newer, currently-cited `outputs/*.md` carry *less* provenance than the older
`past/*.md`.
*Options:* write to a tracked timestamped path by default · keep `build/` and add `--out` plus a
documented archive step (still manual, which is the thing that fails) · refuse to start without a
durable artifact path · normalise the five existing artifacts onto one convention.

**C-9 · residual stale claims the phase-A1 banner does not cover** — slows-work · free-offline · open
`BENCHMARK_RESULTS.md:89-91` claims `daily_max=21` is the minimum feasible and `daily_max=5`
infeasible; measured, 5 is feasible at `--days ≥ 57` and 4 at `≥ 68`, and only 3 is infeasible at
any ceiling. `:18-19` still reads "SMOKE VERIFIED — ready for the full run" six lines below
the banner contradicting it; `:175` still shows opus as `running`. The §4 "reproduce" block
cannot be run (see the hazard box). Dead `--needles` flag in `RUNNING.md:75,91` and
`RUN_RESULTS.md:11`. `RUNNING.md:42` documents a retired one-email-per-turn harness against
`runner.py:412`, and `:3` claims "Everything is reproducible. Same seed in, same result out.",
true of the serve plan and false of a live run. `RUN_RESULTS.md` publishes an 86% headline with
no banner, on a corpus wiped in 2026-06, with a T1–T4 tier vocabulary against
`sb/schema.py:81`'s `T1|T2|T3`. `run.sh:6` names `claude-sonnet-4-6`, which appears nowhere else.
*Options:* extend the banner treatment and correct the specific claims in place · move superseded
docs to `docs/history/` · fix only the mechanically checkable items now and defer prose to phase 4
· add a docs test that fails on a non-existent CLI flag or a 404 link.

### K — corpus authoring (`docs/_repair/K.md`)

Framing: the corpus lints clean (`sb/schema.py:490`) and the oracle scores 167/167 at seeds 42,
1, 99, 2026. **Every defect below is invisible to both checks.** Lint validates shape, the oracle
validates satisfiability, and nothing validates that the rendered email a model reads is
consistent with the answer key it is graded against.

**K-1 · the serve plan drifts past the dates the corpus narrates** —
blocks-measurement · free-offline · open
18 of 167 emails (11%) are delivered on a day *after* a date stated in their own rendered body,
and **11 answer ops demand a calendar object dated before the email that asks for it arrives**
(worst: `World_Cup_Cleat_Launch.json:347`, 14 days in the past). The cause is structural: 143 of
167 emails carry no `date` dependency edge, and lint check #4 (`sb/schema.py:530-549`) only
demands one when an answer references an *anchor* — a body `{dom:13,0m}` paired with an answer
`eq: dom:13,0m` references no anchor and is waved through
(`sb/scheduler.py:76-110`: no date edge → never deadlined → never forced). `_ThisWD.eval`
(`sb/resolver.py:145-148`) computes Monday-of-serve-week plus offset, so `this:FRI` served on a
weekend always resolves into the past, and 49 of 167 emails (29%) are delivered on a weekend. The
five nodes with **zero** date edges own 16 of the 18 past-narrating emails. These emails land in
`grader.py:177`'s `"on the wrong day"` bucket — the one category the brief treats as
unambiguously measuring temporal reasoning.
*Options:* lint that every date-bearing email carries a `date` edge (~143 emails gain edges;
over-constraining is what already makes 19% of seeds infeasible) · scheduler invariant that
`build_plan` fails on a past-resolving answer (would fail today at seed 42) · re-author to
anchor-relative expressions · quarantine the 18 and report them separately.

**K-2 · corpus date-coherence needs a high `daily_max`; the two headline runs used 5** —
blocks-measurement · free-offline · open
The whole grammar is serve-relative by design (`sb/resolver.py:15`), so changing only that lever
**moves 87 of 127 resolved answer dates** and 160 of 167 serve dates, and roughly triples the
defect count:

| `daily_max` | days | prose contradicts its answer | past-dated ops | emails narrating a past date |
|---|---|---|---|---|
| 3 | **infeasible** | — | — | — |
| **5** (opus, sonnet) | 57 | 4 | **11** | **18** |
| 8 | 37 | 2 | 16 | 16 |
| 13 | 21 | 1 | 11 | 17 |
| **21** (documented, haiku) | 16 | 1 | **4** | **6** |
| 30 | 11 | 1 | 2 | 3 |

Read the table honestly: coherence improves **monotonically above `daily_max=13`**, and `30` is
strictly better than `21` on every column (1/2/3 defects against 1/4/6). `21` is the *documented*
lever and the one haiku ran, not the uniquely coherent one — an earlier draft titled this finding
"only coherent at `daily_max=21`", which its own table contradicts. What the lever buys is bought
by compressing the calendar, which trades against V-1.

The sharpest instance: `corpus/nodes/pizza-party.json:121` says *"Could we move the pizza party
to June 9"* with answer `eq: nth:2,TUE,0m`. At `daily_max=21` that resolves to **2026-06-09** and
the prose is correct; at `daily_max=5` the email is served 2026-07-05 and the answer resolves to
**2026-07-14**. The free-typed date is not sloppiness — it is a correct date frozen against a
lever the recorded runs did not use. Separately, 19 of 100 seeds raise `InfeasibleSchedule`
(mechanism: `sb/scheduler.py:85-99`, a deadline computed once and never recomputed, plus `:109`
returning `False` forever once passed).
*Options:* pin `daily_max=21` and re-run there (compresses retrieval span further) · pin
`daily_max=5` and repair the corpus to be coherent at it · make the corpus lever-invariant by
forbidding free-typed dates · stamp levers everywhere and treat cross-lever comparison as invalid
by convention.

**K-3 · prose states a date the answer key contradicts** — distorts-measurement · free-offline · open
Eight emails tell the model one date and grade it against another; unlike K-1 these are not serve
-order artifacts — body and key disagree about *which* date is meant. Gaps of 61, 35, 31, 31, 30,
8, 1 days plus one `by:` window that excludes both readings of "Friday". Two mechanisms: a
free-typed month-day that never traced to a token, and an answer expression **copy-pasted from an
ancestor's emitted expression** rather than referencing the anchor. The cleanest instance:
`Day-of-execution_and_Aftermath.json:15` emits `{!launch_livestream = dom:10,+2m}` and is served
2026-06-10 (anchor = 2026-08-10); `:36` copies the literal `dom:10,+2m` into its answer instead of
`@launch_livestream`, is served 2026-07-22, and resolves to 2026-09-10 — a "green room an hour
before the livestream" scheduled a month after it. Exactly two such copies exist corpus-wide.
**Detector honesty:** the automated relative-phrase detector flagged 9 emails and hand-checking
left 4 true; all five false positives were the same shape (the phrase named a different event than
the graded one).
*Options:* lint bodies for month-name-plus-day outside a token (catches 4 of 8) · lint an answer
expression byte-identical to an ancestor's emit (catches 2, zero false positives) · render-time
cross-check that a date in the body equals the resolved answer (catches all 8, needs a whitelist)
· hand-repair the eight with no rule.

**K-4 · `Innovation-comp` supplies 29% of the emails and ~38% of every model's score** —
distorts-measurement · free-offline · open
One of fifteen nodes contributes **48 of 167 emails** but only 17 of 134 graded ops, has **zero
T3**, and holds **31 of the corpus's 56 no-action emails**. Its share of each model's total score
is 36% / 40% / 37%. Strip the no-action emails and the three models score 35% / 36% / 38% — a
3-point spread on the acting half against the 5-point headline spread. Two nodes no model can do:
`Partnership-with-deeptech-companies` scores 2/10 for **all three**, `World_Cup_Cleat_Launch`
7–8/22; a defect all three hit identically is more likely corpus or grader than capability.
**Stated limit:** this establishes that the node is numerous, easy and homogeneous, not that its
emails are bad. A high no-action fraction is defensible design; what is not defensible is that the
fraction is unstated and unreported.
*Options:* report score split by `has_ops` and by node (no corpus change, no number changes) ·
move surplus no-action emails into `sb.scale`'s filler pool · weight the score by op count ·
add T3 content so the 29% share buys proportionate difficulty.

**K-5 · rendered tokens fuse into prose, render at the wrong grain, or leak authoring notes** —
distorts-measurement · free-offline · open
`resolver.human()` (`sb/resolver.py:368-378`) has exactly one output shape
(`"Monday, June 22nd, 2026"`) and authors wrote prose expecting others. Result: **9** emails where
a full date is fused to an adjacent word ("final order by tomorrow**Friday, June 19th, 2026**" —
genuinely ambiguous, and that email's two readings differ, making it a K-3 case too), **8**
places where a determiner precedes a full weekday-date, **2** naked `{serve}` tokens stamping the
delivery date after a sign-off, and **one leaked authoring note**: `press-tour.json:112` ships
`[insert date → the 13th, +2 months, + add time 11:00 AM-12:00 PM]` verbatim to the model, stating
the answer-key expression in plain English next to the rendered date. That is a scoring-integrity
problem independent of legibility; grep confirms it is the only one.
*Options:* render-time lint on adjacency and bracket notes (needs to run per serve plan) · give
`human()` a short form selected by a token modifier (changes the grammar the webapp TS types
mirror) · hand-fix the 12 affected emails · fix only the leak.

**K-6 · answers the prose does not pin** — distorts-measurement · free-offline · open
Five emails demand an exact calendar day while containing **no temporal cue of any kind** — three
of them `Innovation-comp` `serve+1d` / `serve+2d` / `serve+3d` on bodies whose only urgency cue is
"quick" or "shouldn't take long", with `tolerance: exact_day` (the default, `sb/schema.py:151`).
At the other end, 12 `by:` answers accept a window averaging **19.5 days, max 64** — a 64-day
window on a 57-day run is satisfied by essentially any date. And two `eq:` answers resolve against
an Interval anchor, which `grader.py:78-79` silently reinterprets as "anywhere in that week" while
the prose says "the first day". The corpus is simultaneously too strict and too loose, in ways
that do not cancel.
*Options:* lint that an `eq` answer is traceable to something in the body · cap `by:` windows and
lint the rest · forbid `eq` on an Interval-resolving expression, or make `_matches_value` raise
(a grader change) · convert the five to `within:Nd` tolerance.

**K-7 · `kind` and calendar-plausibility choices** — distorts-measurement · free-offline · open
Four `create` ops declare a `kind` that contradicts what the obligation plainly is — most visibly
a *party* keyed `kind: todo` (`pizza-party.json:22-23`). Because `grader.py:151` filters the pool by
`op.kind` before matching titles, a model that creates the right object of the other kind scores
`(nothing matching created)`, identical in the log to doing nothing. The pizza case compounds:
`_wire_obligations` (`sb/schema.py:393`) makes the sibling `move` inherit `kind: todo`, so a model
that reasonably books the party as an event fails **two of that node's four** graded emails
(`pizza-party.json` grades `end-of-year-pizza-party`, `pizza-place-selection`,
`pizza-order-deadline`, `client-demo-conflict`).
Separately, 14 of 112 `eq` ops land on a Saturday or Sunday, including a public keynote and a
design sign-off, because `dom:N`, `serve+Nd` and anchor arithmetic have no business-day awareness
even though `add_business_days` exists (`sb/resolver.py:80`). **The detector is deliberately
conservative** — 4 of 108 `create` ops flagged, each hand-confirmed — and will have missed
ambiguous cases.
*Options:* re-key the four ops · make the grader kind-tolerant (a substantial contract change,
G's territory) · add an advisory lint heuristic · add a business-day lint or route answers through
`+Nbd`.

**K-8 · anchor, name and metadata hygiene** — slows-work · free-offline · open
**11 of 47 emitted anchors (23%) are referenced by nothing**, including a pair differing only in
capitalisation that holds two dates five days apart for the same party
(`pizza-party.json:149` `{!new_party_date = dom:9,0m}` → 2026-07-09 and `:170`
`{!New_party_date = nth:2,TUE,0m}` → 2026-07-14, while the graded answer is 2026-07-14 and the
free-typed prose says June 9 — three stated party dates in one node).
`_build_emission_map` (`sb/schema.py:451-460`) raises only on exact-name collisions. Unused emits
are not free: they reserve a name corpus-wide. Nine obligation names carry edge whitespace, and
because `match` defaults to `[name]` (`sb/schema.py:155`) a trailing-space name becomes a
**match keyword containing a trailing space**. 13 duplicated subject lines (`"Reveal event date
and venue"` ×3) degrade the retrieval surface `list_new_emails` presents; 4 emails have an empty
`from`.
*Options:* lint unused anchors, case-pairs and edge whitespace (fails the corpus today at 11+1+9
sites) · strip whitespace at parse time and warn (hides the authoring error) · lint or auto-suffix
duplicate subjects (changes what the model reads) · hand-clean with no rules.

### V — construct validity (`docs/_repair/V.md`)

**V-1 · the corpus is two orders of magnitude too small to push any fact out of context** —
blocks-measurement · free-offline · open
`sb/span.py:4-6` states the claim: a large span means "the fact has likely scrolled out of context
… so the model must `search_inbox` to recover it." The **entire authored corpus is 43,474
characters (~10.9k tokens)**, the mean authored body is **227 characters**, and the largest gap
between a needle's setup and its payoff is 18,389 body chars (~4.6k tokens), rising to ~5.4k
tokens once subject and sender are counted over the same 83-email window, and higher again with
per-email JSON overhead. (An earlier draft said ~7.8k; that figure required an unstated ~10k-char
JSON envelope. The conclusion is untouched either way.) Every model in the roster has ≥200k context.
**No authored fact has ever scrolled out of context in any recorded run**, and `search_inbox` has
never been necessary to answer a single email.
*Options:* scale the corpus with filler until the gap exceeds a real window and re-run (gated on
V-5, and paid) · rescope the claim to multi-day dependency reasoning within context (free,
immediately honest, abandons the differentiating claim) · force retrieval mechanically by eliding
older turns (makes span a controlled variable at any corpus size; changes the harness contract) ·
keep the corpus and publish the null result.

**V-2 · the search-rate figure is real for opus and unmeasurable for sonnet and haiku** —
blocks-measurement · free-offline · open
The day prompt carries no bodies — `runner.py:418-424` POSTs bodies to the store and `:426-429`
sends only the date — so the only ways to read a body are `get_email` (`store_app.py:225-230`) and
`search_inbox` (`:192-213`). A model acting on N emails must therefore call `get_email` at least N
times. Opus retained **166 `get_email` for 167 emails** with one deficit day out of 57 (day 1,
the one day it searched). Sonnet retained exactly one on each of 57 days (57 total) and haiku 40
over 16 days. Additionally, the signal is pinned to zero by construction: a needle's payoff must
be served after its emitter, so no needle can fall on day 1, and day 1 is the only day either
model searched — running the retrieval analysis on either full run gives `searched% 0%` in
**every** span bin.
*Options:* fix `_parse_stream` to append and re-run (does not recover the existing logs) ·
instrument retrieval server-side (driver-independent, needs a persisted artifact) · make the turn
boundary one email instead of one day (exact attribution; multiplies turn count and spend) · drop
`searched` from the reported analysis.

**V-3 · 38.3% of the score is available to a model that never calls a tool** —
distorts-measurement · free-offline · open
A null model taking no action scores **64/167 = 38.3%**: all 56 no-action emails pass
(`grader.py:186-195`) and all 8 cancel-only emails pass, because `grader.py:155-158` passes a
`cancel` when nothing matching is on the calendar — trivially true if the object was never
created. Against that floor:

| run | raw | on the 64 a null model passes | on the 103 requiring action | normalised above floor |
|---|---|---|---|---|
| opus | 90/167 = 53.9% | 56/64 | **34/103 = 33.0%** | 25.2% |
| sonnet | 91/167 = 54.5% | 55/64 | **36/103 = 35.0%** | 26.2% |
| haiku | 98/167 = 58.7% | 60/64 | **38/103 = 36.9%** | 33.0% |

Normalised, the spread widens from 5 points to 7.8, still with haiku on top, and haiku's advantage
is concentrated in the free points (60/64 vs 56/64).

**The system prompt instructs models to take the floor.** `runner.py:101` reads *"Over-acting is
the most common mistake; when in doubt, do nothing."* V-3 establishes the 38.3% is *available*;
that line makes claiming it the instructed strategy. No section owned this — the verifier found it
while checking V-3 — and it belongs to whatever decision V-3 becomes, because it means the floor
is not an artifact a model has to discover, it is advice.
*Options:* report the floor alongside every score and the actionable subset separately (free, one
line, changes nothing) · make `cancel` falsifiable by requiring the object to have existed (a
grader change, removes 8 free points) · rebalance the no-action fraction · score by op.

**V-4 · 190 op-level judgements are compressed into 167 binary points** —
distorts-measurement · free-offline · open
`grader.py:197-198` sets `passed = all(...)` and `runner.py:519` sums booleans, so a 4-op email is
worth the same one point as an FYI, and 3-of-4 scores identically to 0-of-4. Ops per email:
`{0: 56, 1: 94, 2: 12, 3: 4, 4: 1}` = 134 ops + 56 no-action checks = **190 judgements collapsed
to 167 points**, confirmed by each log containing exactly 190 `why` lines. `grader.py:202` keeps
only the *first* failing reason in `headline`. The taxonomy survives in the printed log — the
§2.3 tally reproduces exactly — but not in the metric, and not in any machine artifact.
*Options:* keep the binary headline and additionally emit per-op counts and reason buckets ·
switch the denominator to 190 ops (breaks comparability with all four runs) · report both · split
the reported score by answer shape (n = 56/94/17, the last too small to be stable).

**V-5 · `sb/scale.py` is unusable as run and self-defeating where it works** —
blocks-measurement · free-offline · open
The one instrument that could establish V-1's claim **fails non-monotonically**: `--filler` 30, 60
and 200 raise `InfeasibleSchedule` while 90, 120 and 150 succeed, and at `--days 400`, 175 succeeds
with a *lower* needle span than 150. The error's own advice is wrong — `sb/scale.py:113-114` says
to raise `--days` when `--days 300` was already supplied and the real cause is over-constrained
serve windows. Where it works it defeats itself: every filler email is a graded no-action email
(`sb/scale.py:80`: `"answer": {"ops": []}`) and the runner grades everything in the plan with no
filter (`runner.py:503-511`), so the do-nothing floor rises **38.3% → 59.9% / 64.1% / 67.5%** at
filler 90 / 120 / 150 — above every score any real model has recorded. A run at `--filler 120`
would produce a number that looks like a benchmark score and means "the model correctly ignored
176 newsletters."
*Options:* tag and exclude filler from grading (restores the floor; filler stops testing
over-action) · weight or subsample filler · fix the scheduler's serve-window handling first ·
abandon volume scaling for explicit context eviction.

**V-6 · `tier` is loaded and never read; the by-tier report hangs off a file nothing writes** —
blocks-measurement · free-offline · open
`TIER_LIST.md:175-181` asks for a score-by-tier report and states "The score that matters is T3".
Every email is tagged (50 T1 / 67 T2 / 50 T3). `Email.tier` is declared at `sb/schema.py:81` and
populated at `:230`, and **`grep -rn "\.tier\b" --include="*.py"` returns zero hits** — the field
is write-only. `sb/analyze.py:94-105` instead looks for `reasoning_tier` in
`<corpus>/needles.json`, a file that exists nowhere and that no code writes, so the tier column
always prints `untagged`. The `100+` span bin (`analyze.py:29`) is structurally unreachable on the
authored corpus (max span 83). **T3 accuracy has never been computed for any run.**
*Options:* read `corpus.emails[eid].tier` directly and delete the `needles.json` branch · emit a
`needles.json` manifest from `sb.scale` · report tier accuracy from the runner once O lands · drop
the tier concept from tooling.

**V-7 · the authored tier gradient does not exist on the axis it is defined by** —
distorts-measurement · free-offline · open
`TIER_LIST.md:23-27,191` define T1→T2→T3 primarily by retrieval distance. Measured, **T3's needles
are closer than T2's on both axes** (email-span 31.1 vs 33.3; day-span 10.4 vs 11.8 — but on
n=6 T3 needles against n=18 T2, so the comparison is directional, not significant), 32 of the 50
T3 emails have no cross-email anchor reference in their answer key at all, and 13 of the 50 are
pure no-action, which `TIER_LIST.md:57-59` assigns to T1. `:166-168` requires every T3 to stack
≥2 hardeners; only 18 of 50 carry even one measurable one. **Stated limit:** Hardeners B and C are
not machine-visible, so this is a lower bound on hardener count, not a refutation of the other two.
*Options:* re-derive tiers from measured span and report both · keep author tiers, add a measured
axis, report the crosstab · re-author the T3 set to satisfy Hardener A (gated on V-5) · retire
Hardener A from the definition (then V-1's claim has to go too).

**V-8 · the span axis is reconstructed from levers the artifact never recorded** —
distorts-measurement · free-offline · open
`sb/analyze.py:88-91` rebuilds the plan from CLI flags and computes span from that
reconstruction, because the run saved nothing but text. Same opus log, `--daily-max 5` vs `21`:
bins move from 19/5/0 to 13/10/1 and the top bin appears only in one of them, while overall needle
accuracy is 42% either way — so nothing detects the mismatch. Reported as tool output only, not as
a result: both recorded runs show accuracy *rising* with span (opus 37%→60%, sonnet 42%→60%), the
opposite of the hypothesis in `analyze.py:2-4`, at n=19 and n=5 with bin membership a function of
the guessed lever.
*Options:* have the runner emit each email's span alongside its verdict (needs O) · parse and
enforce the stamp (leaves the pre-phase-0 logs unanalysable) · pin one canonical lever set and drop
the flags · document required flags per artifact and keep the silent-wrong-answer path open.

---

## Cross-category resolutions

Three places where sections disagree. Each is called, with the reasoning.

### 1. The Haiku-above-Opus anomaly: three mechanisms, one confound, and the test that separates them

G, A and K each offer an explanation. They are not competing hypotheses about one quantity —
**they explain different quantities**, and only one of the three is contested.

First, the quantity to be explained. Haiku's 8-point lead over opus decomposes (A-1, K-4, V-3) as
**5 points of no-action** (56/56 vs 51/56) and **3 points of acting emails** (42/111 vs 39/111).
Normalised above V-3's 38.3% null floor the lead widens to 7.8 points and is *more* concentrated
in the free points (haiku 60/64 vs opus 56/64).

**G explains the band, not the ordering.** The shipped `sb.engine` + `sb.grader`, driven by an
agent that is scheduling-perfect in every respect and titles calendar entries after the email
subject, scores **92/167 = 55%** — inside the 54–59% band of all three real runs (G-4, P3). This
is the single best-supported measurement in the fan-out: shipped code, shipped corpus, no live
run, no inference. It establishes that the observed band is fully reachable with zero scheduling
errors. It says nothing about why haiku sits 5 points higher than opus, because no-action grading
never touches `_title_hit` at all.

**K explains why haiku was asked different questions, and this is the decisive confound.** The
answer-key grammar is serve-relative (`sb/resolver.py:15`), so `daily_max` alone moves **87 of 127
resolved answer dates**. Haiku ran at 21 — the lever the corpus's prose was evidently authored
against and the lever `BENCHMARK_RESULTS.md` §1 pins — where past-dated ops fall from 11 to 4,
emails narrating a past date from 18 to 6, and prose-contradicts-answer from 4 to 1 (K-2). **Haiku
was graded against a materially more coherent answer key on the same 167 emails.** This does not
prove the anomaly is a corpus artifact (haiku also saw a different plan, different context
lengths, and 16 turns rather than 57), but it removes "same benchmark, different model" as a safe
reading, and it is the mechanism with the clearest causal path to the 3-point acting gap.

**A's rank inversion is real as a sensitivity result and must not be cited as evidence about
attribution, because it is confounded by the same lever.** A-1's day-scoped counterfactual gives
opus 90→55, sonnet 91→57, haiku 98→43. The confound is **structural and model-independent**, not
a fact about how these three models happened to behave. Under a day-scoped rule a no-action email
can only survive if *nothing* was created that day, so the ceiling is fixed by the serve plan
alone — count the no-action emails sharing their day with no ops-carrying email:

| `daily_max` | days | no-action emails alone on their day | ceiling on day-scoped no-action passes |
|---|---|---|---|
| **5** (opus, sonnet) | 57 | **10 / 56** | 10 |
| 8 | 37 | 3 / 56 | 3 |
| 13 | 21 | 0 / 56 | 0 |
| **21** (haiku) | 16 | **1 / 56** | 1 |
| 30 | 11 | 0 / 56 | 0 |

Haiku's day-scoped score is capped **nine points below** opus's and sonnet's before any model
behaviour is considered. (An earlier draft argued this correlationally, from where each model's
passes happened to land; the ceiling argument replaces it and reaches the same conclusion causally.
It also survives the trace-loss caveat: haiku has the lossiest trace and therefore the most room
for a *spurious* surviving pass, yet has the fewest, because the ceiling binds.)
**Within the lever-matched pair the ordering is preserved** — sonnet leads opus both
before (91 > 90) and after (57 > 55). Only the odd lever out inverts. The honest statement of A-1
is that the abstain check is 56–57% of every score and its only input is a model-supplied id;
the "the ranking inverts" headline is a lever effect wearing an attribution label. A itself flags
the counterfactual as deliberately extreme and as measuring sensitivity, not correctness.

**Call.** All three are real; they are not equally supported and they are not about the same
thing. G-4 sets the level and is the strongest evidence in the fan-out. K-2 makes the haiku number
incomparable to the opus and sonnet numbers and is the reason the anomaly should not be quoted at
all until re-measured. A-1's contribution is the weight of the abstain check, not the inversion.

**What would separate them, in order of cost:**
1. *Free, not yet run.* G-4's own open question: check whether a P2/P3 title-policy oracle fails
   the **same obligations** the real models fail. High overlap implicates G's mechanism; low
   overlap implicates the models. Possible offline today against the recorded logs.
2. *Free.* Re-run K-2's defect sweep at every feasible lever and check whether free-typed-date
   agreement peaks at 21, confirming it was the authoring lever. K measured 5 and 21 only.
3. *Paid, phase 7, and the only clean test.* Run at least two models at **both** levers. That is
   the sole way to separate lever from model, and it is why phase 7 must pin one config and run
   the roster inside it rather than comparing across the historical three.

Until (3), **the correct statement is that opus and sonnet are comparable to each other and
neither is comparable to haiku.**

### 2. The tool trace: what is lost, and exactly which conclusions survive

O-2 reports opus losing `create_*` calls on 18 of 57 days. V-2 reports opus's `get_email` complete
on 56 of 57 days (166 retained for 167 emails) and uses that to establish opus's low
`search_inbox` count as a real behavioural fact. Both cannot be maximally true.

**The mechanism, stated once.** `sb/live/runner.py:183` is `seen[msg.get("id", id(ev))] = msg`,
and `:187-194` iterates only the surviving `msg`. Loss occurs in exactly two circumstances:
(a) one model message id spans multiple assistant events, i.e. the model emitted a batch of
parallel tool-call blocks — every block but the last is overwritten; (b) the message carries no id
or a null id, where the fallback `id(ev)` is the CPython address of a dict rebound each iteration
and the allocator reuses addresses (measured: 4 of 6 retained with no id, 1 of 6 with null). The
`codex` driver appends instead (`runner.py:355-357`), so trace fidelity differs by driver. **Loss
is therefore a function of how a model chunks its tool calls, not a constant.**

**What survives.**

- **V-2's opus conclusion stands, with its argument tightened.** Opus's per-day `get_email`
  histogram is `{1:11, 2:13, 3:12, 4:12, 5:9}`, tracking the day's batch size exactly, and totals
  166 for 167 emails. Retained can never *exceed* actual, so the trace **meets its lower bound
  on 56 of 57 days** — a tight lower bound, and the strongest statement the data supports. It is
  not proof that no `get_email` was dropped, and the earlier draft's further step — that count
  equality *demonstrates* opus serialised its tool calls, one content block per assistant
  message — does not follow either; O's own synthetic case 2 refutes it. A message emitting
  `[search_inbox, get_email]` would preserve the `get_email` count while losing the search, so
  a hidden search is bounded, not excluded. Day 1's line,
  which shows four separate retained `get_email` entries interleaved with three retained
  `search_inbox` entries, is direct evidence of serialisation.
- **Sonnet and haiku are demonstrably lossy.** Sonnet's `{1: 57}` — exactly one `get_email` on
  every day regardless of 1 to 5 emails arriving — is impossible for a complete trace given the
  hard lower bound of one `get_email` per email read. Haiku: 40 retained for 167 emails.
- **O-2's mechanism, both bugs, the driver asymmetry, and the `analyze.py` dependency all stand
  untouched.**

**What does not survive.** **O-2's claim that opus's trace is lossy is withdrawn.** It rests on a
per-day comparison of `create_*` calls against titles "appearing for the first time" in an `actual`
field. But `actual` is drawn from the node's **cumulative** pool (`grader.py:151-152`), so an
object created on day 10 in a node with no graded op until day 20 first surfaces on day 20. First
appearance in `actual` is an *upper bound* on creation day, not the creation day, which
manufactures deficits on some days and hides them on others. O's own aggregate for opus — 89
retained `create_*` calls against 71 distinct titles — shows **no proof of loss**, and O says so in
its own table. The same aggregate test is decisive against sonnet (41 vs 70) and haiku (18 vs 68),
so the measure is sound in aggregate and unsound per day.

**Consequences for other findings.**
- Brief §2.8's search counts are a **measurement** for opus (1 of 57 days) and a **lower bound**
  for sonnet and haiku. Phrase them that way everywhere.
- O-2's "41–74% `create_*` drop rate on those same runs" applies to sonnet and haiku, not opus.
- V-1's central claim is untouched by any of this: it is a measurement of the corpus
  (~10.9k tokens, ~5.4k max needle gap), not of the trace.
- A's use of the tools line to rule out whole-day inaction (see resolution 3) is sound in the
  direction it is used: loss can hide a `create_*` that happened, never invent one that did not.
- The one still-open question is whether the current CLI emits one `assistant` event per content
  block. Both O and V infer it from the logs; neither confirmed it against a captured stream.
  Capturing one turn of raw `stream-json` settles it and is the cheapest possible live check.

### 3. §4.3 — the split of the 52 "nothing matching created" cases

G partitions the bucket with three offline signals; A independently constrains one arm; O argues
the whole exercise is unmeasurable. These reconcile into one statement.

**The measured partition (G, all 52/47/52 cases mapped, 0 unmappable):**

| | opus (52) | sonnet (47) | haiku (52) |
|---|---|---|---|
| S1 keyword appears nowhere in the node's rendered mail | 27 (52%) | 28 (60%) | 29 (56%) |
| unexplained (under-action or paraphrase) | 16 (31%) | 13 (28%) | 13 (25%) |
| S3 same-keyword object visible under another node | 8 (15%) | 6 (13%) | 8 (15%) |
| S2 proven kind mismatch | 1 (2%) | 0 | 2 (4%) |

**A's independent constraint on the residual:** whole-day inaction is ruled out for **49 of opus's
52** and 42 of sonnet's 47, which fall on days whose tools line shows at least one
`create_event`/`create_todo`. Because the trace is lossy this is a lower bound in the safe
direction. It does not prove the model acted on *that* email, only that it was not globally idle.

**A's negative result on the attribution arm:** there is **no positive evidence** for cross-node
misattribution in any run — every object in every duplicate-failure `actual` field is topically
native to the node it is listed under, and the one confirmed misattribution (A-6) was *within* a
node. G's S3 signal (13–15%) is explicitly a "worth checking" heuristic, not proof: a shared
keyword such as `atlas` or `design` can legitimately appear in another storyline's objects.

**O's correction to the framing, which is right:** brief §4.3 should read **"unmeasurable from any
surviving artifact"**, not "unmeasured". `grader.py:152,168` renders only keyword-matched objects
and the store was discarded (`runner.py:513-516`), so no surviving file can attribute an
individual case. G's partition is not a per-case attribution and does not claim to be.

**Merged statement.** §4.3 is **bounded, not closed**:
- The bucket is **majority a grader-plus-corpus discoverability failure** (S1, 52–60%): the grader
  demanded a string the model had no way to know. Read as an **upper bound on model fault**, not
  proof the model created something wrongly titled.
- **Genuine under-action is at most 31% / 28% / 25% of the bucket**, and within that residual,
  global idleness accounts for at most 3 of opus's 52 (A). So the residual is per-email paraphrase
  or per-email inaction on an otherwise active day.
- **Kind mismatch is the smallest arm by an order of magnitude** (0–4% proven). K-7 names four
  corpus-caused candidates, which converts it from hypothetical to bounded and small. The brief
  listed it as one of four *equal* candidates; it is not.
- **The attribution arm has no positive evidence and cannot be refuted from logs.** It stays open
  as an exposure argument (A-5) until an instrumented run measures it.
- Closing it per-case requires phase 6 + phase 7. Nothing on disk can do it.

---

## Overlap map

Built from the "Overlaps with:" lines of all six sections. Read as: work on the left cannot be
scoped without deciding the right.

| finding | overlaps |
|---|---|
| G-1 | G-3, G-4, G-5, G-8; K-1, K-8; A (attribution as an alternative identity channel) |
| G-2 | G-1, G-3, G-5, G-6, G-7; C-3; K-8 |
| G-3 | G-1, G-2, G-9; O-5 |
| G-4 | G-1, G-2, G-5, G-7; C-3; V-1 |
| G-5 | G-1, G-2, G-7, G-8; C-3; K-8 |
| G-6 | G-2, G-7; A-5; K-4; O-3, O-5 |
| G-7 | G-2, G-5, G-6; A-3; K-7 |
| G-8 | G-1, G-5; K-8 |
| G-9 | G-1, G-3, G-6; A-5; O-5 |
| G-10 | all of G; O-5 |
| A-1 | G-*; K-2, K-4; V-3, V-4; C-1 |
| A-2 | G-6; O-1, O-3 |
| A-3 | A-4, A-6; G-6; O-3, O-9 |
| A-4 | A-3, A-5; C-1; O-1 |
| A-5 | A-3, A-4; G-6, G-9; O-3, O-5 |
| A-6 | A-3, A-4; G-6, G-7; K-4 |
| O-1 | C-1, C-6, C-8; O-3, O-5 |
| O-2 | V-2; O-4, O-6; C-2 |
| O-3 | O-1; A-5; G-6; C-1, C-5; K-2 |
| O-4 | O-2; V-1, V-2; A-4 |
| O-5 | G-1, G-9; A-5; O-1, O-3; §4.3 |
| O-6 | O-2, O-4; V-2, V-6; A-1 |
| O-7 | A-3; G-2; C-7; O-1 |
| O-8 | C-1, C-5; M-5; O-1 |
| O-9 | A-3; G-6; O-5 |
| C-1 | O-1, O-3; K-1…K-8 (phase 5 triggers it); M-2 |
| C-2 | C-1, C-4, C-5; K-2; O-1; V-8 |
| C-3 | G-2, G-5; K-2, K-6; C-2 |
| C-4 | O-1, O-6; V-6, V-8; C-2 |
| C-5 | C-1, C-2; O-1, O-3 |
| C-6 | C-1, C-8; M-5 |
| C-7 | M-2, M-3; O-1, O-8 |
| C-8 | C-1, C-6; O-1 |
| C-9 | C-1, C-5; V-1, V-6, V-7; M-4 |
| K-1 | K-2, K-3, K-7; G-9; V-4 |
| K-2 | K-1, K-3; C-1, C-2, C-3; V-1; A-1 |
| K-3 | K-1, K-2, K-5, K-6; G-9 |
| K-4 | A-1, A-3; G-1; V-3, V-4; C-1 |
| K-5 | K-3, K-8; V-1 |
| K-6 | K-3; G-2, G-7; V-4 |
| K-7 | G-1, G-9 (kind filter is `grader.py:151`); K-1; A-5 |
| K-8 | G-1, G-5, G-8; K-5 |
| V-1 | V-2, V-5, V-7; K-2 |
| V-2 | O-2, O-4, O-6; V-1, V-8 |
| V-3 | G-7; A-1; V-4; K-4 |
| V-4 | G-9; O-5; V-3 |
| V-5 | V-1, V-3; K-1, K-2 |
| V-6 | O-6; C-4; V-7 |
| V-7 | V-1, V-6; K-4 |
| V-8 | C-1, C-2, C-4; O-1; V-2 |

**Densest couplings, i.e. what cannot be decided in isolation:**
- **G-2 ↔ G-6 ↔ G-7 ↔ C-3.** The exactly-one rule, the cumulative pool, cancel-by-absence and the
  lever-dependent oracle failure are one mechanism seen four ways. A best-match assignment (G-2
  opt 1) resolves all four; anything less resolves none of them cleanly.
- **A-1 ↔ K-2 ↔ K-4 ↔ V-3.** The no-action fraction, the lever confound, the single-node
  concentration and the null floor are one number (56–57% of every score) described four ways.
- **O-1 ↔ O-3 ↔ O-5 ↔ C-1.** What a run leaves behind, and what would have to be captured for
  phases 2 and 3 to be checkable against real behaviour.
- **V-1 ↔ V-5 ↔ V-7 ↔ K-2.** The retrieval claim, the instrument meant to establish it, the tier
  construct that encodes it, and the lever that determines the corpus's coherence.

---

## Corrections to `docs/benchmark-repair-evidence.md`

**The brief is not edited.** It is the historical input to this fan-out and the verifier compares
against it. Twenty corrections, all found independently by two or more agents where marked (†).

**Code anchors (all confirmed against the working tree at `67b3005`):**

| brief says | what is actually there | correct anchor |
|---|---|---|
| §2.6 `runner.py:472-475` "grades a day's objects by splitting on `email_id`" † | the retry error / usage-pause block | `runner.py:498-510` |
| §3 `runner.py:394-477` "day loop, attribution split" † | `order = [...]` flattening | day loop `runner.py:412-512`; split `:498-510` |
| §2.4 / §3 `grader.py:163-165` "`pool` is all objects of that kind" † | `matched` / `count_ok` / `passed` | `pool` is `grader.py:151`; `title_set` `:152`; `count_ok` `:164` |
| §3 `runner.py:180` "lossy tool parse" | `elif t == "assistant":` | the overwrite is `runner.py:183` |
| §2.12 `runner.py:308-310` "`--permission-mode bypassPermissions`" | the `build_cmd` signature | `runner.py:320-322` |
| §2.5 `oracle.py:51` "oracle titling" | `for op in email.answer.ops:` | `oracle.py:52` |

`store_app.py:86-92`, `analyze.py:25`, `grader.py:68-70`, `schema.py:155`, `span.py:26-41` and
`scale.py:67-96` all check out.

**Substantive:**

1. **§2.1 is incomplete: there is a fourth complete run log.** † `past/claude-sonnet-4-5.md` is
   1047 lines with header `Corpus: 176 emails (sha 809d389794dd79a9) · seed 42 · days 30 ·
   daily_max 21`, `SCORE 102/176 (58%)`, `PASS 102 · FAIL 74 · ERROR 0 · search_inbox 1`, and 176
   parseable per-email verdicts. It is a second corpus showing the same grader signature (14
   collide failures, 0 of them same-title, the same `atlas`/`design`/`launch`/`dynamics`
   collisions), which is what makes the signature corpus-independent.
2. **§2.1's "two evidence links point at files that do not exist" is right but incomplete.** † The
   *links* (`outputs/claude-*.md`) are dead and are the **only** broken markdown links anywhere in
   the tracked doc surface. But the sonnet artifact survives at `past/claude-sonnet-4-5.md`; only
   the haiku 84/176 artifact is genuinely gone, overwritten by a later run of the same model at the
   same model-id filename.
3. **§2.3's "in all 52 opus cases the `actual` field reads `(nothing matching created)`" is a
   tautology, not an observation.** `grader.py:168` writes that string under exactly the condition
   (`not title_set`) that `:172-173` uses to write the corresponding `why`. The sentence that
   follows it — that the log cannot distinguish the causes — is correct.
4. **§2.6's inference is unsound, not just its anchor.** "Those warnings fired once, so this is a
   theoretical hole" cannot be supported by a monitor that has no branch for a same-day sibling id
   (`store_app.py:86-92`); the sibling case produces zero warnings by construction. And the one
   warning that did fire is an observed, score-changing instance of a *different* escape route
   (A-4, A-6).
5. **§2.7's quoted day-1 trace is attributed to the wrong artifact and is missing its first
   entry.** † The quoted line is `past/claude-sonnet-4-5.md:18` (the retired 176-email corpus, 8
   emails), not `outputs/sonnet.md:8` (5 emails), and it begins `ToolSearch, list_new_emails, …`.
6. **§2.7's "Same harness, so those counts are not comparable" is the wrong inference.** The counts
   are directly comparable and are the most useful diagnostic available: they differ because opus
   serialised its tool calls and sonnet batched them.
7. **§2.7 understates the scope in one direction and overstates it in another.** The trace is lossy
   for sonnet and haiku, not for opus — see resolution 2 above.
8. **§2.8's search counts are a measurement for opus and lower bounds for sonnet and haiku.** The
   offline span measurement (mean 31.6, max 83) is independent of the trace and stands.
9. **§2.5 "the oracle cannot detect this class of bug" is overstated.** The oracle *is* a partial
   detector for between-obligation collisions: at `(2, 5, 7)` it scores 166/167 with
   `found 2 matching, expected exactly 1`. It is blind to keywords a real model would never
   produce, which is the narrower and correct claim.
10. **§2.10 / the register's "every answer key is satisfiable" is true only at `(1, 5, 7)`.** 12 of
    93 feasible lever settings score 166/167.
11. **§2.10: pizza-party has four unreferenced emitted anchors, not three** — `@ordering_date`,
    `@pizza_decision`, `@new_party_date`, `@New_party_date` (the last two are a case pair holding
    dates five days apart).
12. **§2.10: `by tomorrow{!ordering_date = this:FRI}` renders "June 19th, 2026" at the default
    config, not June 12th.** The June 12 rendering corresponds to a configuration the brief does
    not state.
13. **§2.10: `{serve }` is a valid token, there are two of them, and both sit at the end of a
    body**, not "a stray token rendering a date mid-sentence" (`resolver.py:239` accepts bare
    `serve`; `pizza-party.json:14`, `Sponsoring-Marathon.json:15`).
14. **§2.10's framing of the free-typed dates is right about syntax and wrong about cause.**
    `pizza-party.json:121`'s "June 9" is exactly what `nth:2,TUE,0m` resolves to at
    `daily_max=21`. These are correct dates frozen against the documented lever, not invented ones.
15. **§2.1's levers column is under-specified for haiku.** `daily_max=21` is independently
    confirmable from the log's serve plan; `urgency_horizon` is not — **seven** values (1 through
    7) reproduce it exactly. The brief's only source is a hand-typed header whose corpus-sha field
    is unverifiable (C-5).
16. **§2.2's forensics reproduce with a small methodology difference:** em-dash share 61%/0%/0%
    against the brief's 64%/0%/0%; mean title length 47/36/40 against 46/36/40; opus↔sonnet verdict
    agreement 148/167 = 88.6% exactly. Same conclusion.
17. **§2.9's "`analyze.py` never reads `email.tier`" is confirmed and sharpened:** it reads a
    *different key* (`reasoning_tier`) from a *file that does not exist* (`<corpus>/needles.json`),
    which nothing in the repo writes.
18. **§4.1's arithmetic should be restated.** 135/167 outcomes are **model-invariant**, which is not
    the same as **grader-determined** — three models could genuinely all fail a hard obligation.
    The defensible causal subset is the 57 collide failures, the ≥7-per-run collide failures where a
    correct object sat on the correct day, the 99 pooled op-details whose keyword is undiscoverable
    and which pass at 8%, and the 25 description-driven collisions the shipped oracle suffers at
    100% scheduling accuracy.
19. **§4.4 is closed as far as this repository goes.** Across every commit on every ref including
    `origin/backups`, exactly four files have ever contained a `SCORE` line and none scores 51%.
    `git fsck --lost-found` shows no dangling objects, `git stash list` is empty, no `.log` exists
    on disk. Only an off-repo copy could exist. C-8 identifies the mechanism.
20. **§4.5 is established with a stronger statistic than verdict agreement.** Across the three
    167-corpus runs, 77/167 emails (46%) are unanimous passes, 58/167 (35%) unanimous fails, and
    only **32/167 (19%) discriminate between the models at all**. At op level 95/134 (71%) are 0/3
    or 3/3, and 62 ops are failed by every model. On the 111 action emails the three score 39, 40,
    42 and 58 of the 111 are unanimous failures. **The benchmark's entire discriminating range is
    32 emails**, and the observed 8-point spread sits inside a band that is 81% pre-determined.

**Brief claims now upgraded from "not established":** §4.2 is bounded by V-3's 38.3% null floor
(and any hand-grade should sample the actionable 103, not the 167); §4.3 is bounded, see resolution
3; §4.4 closed, see above; §4.5 established, see above. **§4.1 stands as written** apart from
correction 18.

---

## Open questions no category owns

- **The isolated cwd and the answer key.** Brief §2.12 says the model has all built-in CLI tools
  alongside the MCP tools, and that the temp cwd does not block `Read`/`Bash` from reaching
  `corpus/nodes/*.json`. The code's own comment at `runner.py:318-319` claims the opposite —
  "bypassPermissions + strict-mcp-config already scope tools to only our secretary MCP server" —
  while `--strict-mcp-config` (`:322`) scopes MCP servers, not built-ins, and `ToolSearch` appears
  in every run log. *Synthesizer-added; no category agent owns this.* No log shows any model
  reading the corpus, and G notes independently that a model which had would have written
  `breifing` and `Team_pizza_party` into its titles. Worth one bounded check before phase 7.
- **Whether the CLI emits one `assistant` event per content block.** Both O and V infer it; neither
  confirmed it. One captured turn of raw `stream-json` settles it (resolution 2).
- **Whether `codex` behaves comparably at all.** `codex` is not installed on this machine, so
  `_parse_codex`'s resolved-model path (C-7), its cost fields (O-8) and its non-deduping trace
  (O-2) are unexercised for half the documented roster.
- **Whether phase 5's corpus repair should target `daily_max=5` or `21`** (K-2 opt 1 vs 2). This is
  the single decision that most changes the size of the phase-5 authoring pass, and it depends on
  V-1's construct decision. It belongs to phase 4.

---

## Open question inherited from before phase 0

The reported score was "about 51%". No committed log shows 51% — `outputs/opus.md` is 54%,
`outputs/sonnet.md` 54%, `past/claude-haiku-4-5.md` 59%, `past/claude-sonnet-4-5.md` 58%. So at
least one run exists whose artifact was never saved. If that log survives anywhere it is worth
recovering: its header states the model the runner *thought* it was using, which would confirm M-1
directly rather than by inference.

**Phase A update (C-8):** git archaeology is exhausted. Across every commit on every ref including
`origin/backups`, exactly four files have ever contained a `SCORE` line; `git fsck --lost-found`
reports no dangling objects; `git stash list` is empty; no `.log` file exists on disk. The only
remaining possibility is an off-repo copy on this machine (an older working copy, Time Machine, a
`~/Downloads` copy). The mechanism that lost it is C-8: every documented run command redirects into
gitignored `build/`.

---

## Changelog

- **2026-08-18** — Phase 1a. The capture slice is **applied**, not yet verified.

  `sb/live/runner.py` gains `--out DIR`. Per day it writes the store state it was
  already fetching and discarding, which is every object the model created with title,
  description and attribution — including the ones matching no answer-key keyword, which
  the printed log never renders (O-5). Plus the raw CLI stream per day, so a later fix to
  the trace parser (O-2) can be applied retroactively to runs already recorded. Plus a
  manifest carrying the certified served model, levers, seed, serve plan and a corpus hash
  whose algorithm **exists in this repo** and is reproducible (`_corpus_hash`), unlike the
  one C-5 flags in the historical headers.

  New `sb/regrade.py` re-scores a capture offline with no model and no store:
  `python -m sb.regrade <dir>`. This is what O-1 was blocking.

  **Nothing touches the model-facing tool surface**, so a captured run stays directly
  comparable to the uncaptured historical ones. `store_app.py`, `grader.py`, `mcp_app.py`
  and the corpus are unchanged.

  **Named artifact:** `sb/tests/test_capture_regrade.py`, 3 tests. It simulates the day
  loop against a store state in the exact `/state` shape, writes a capture, re-grades it
  offline and asserts the score is identical — at two title policies, answer-key (153/167)
  and email-subject (90/167 = 54%). That 54% independently reproduces G-4's central result
  through a different code path. Suite: 62 → 65 passing.

  **Why `applied` and not `verified`:** the test simulates the day loop rather than calling
  `run()`, so the capture *format* and the *offline re-grade* are proven while the wiring
  inside the live loop is not. That is exactly what phase 1b's bounded smoke is for. No
  finding moves to `verified` until a real run writes a capture that re-grades to its own
  printed score.


- **2026-08-17** — Phase 0. Fixed E-1 (venv on 3.13), E-2 (`mcp<2.0.0`), E-3 (MCP stderr),
  M-1 (`run.sh` argument forwarding), M-2 (resolved-model reporting + config stamp),
  M-3 (per-turn drift detection). Closed M-4. Opened M-5. Confirmed corpus oracle 100%.
  62 unit tests pass. Confirmed corpus oracle 100%.

  Live smokes, all launched via the previously-broken `./run.sh live --model X` form:

  | smoke | requested | served | result |
  |---|---|---|---|
  | 3 emails, 1 turn | `claude-sonnet-4-5` | `claude-sonnet-4-5-20250929` ✓ | 3/3 |
  | 3 emails, 1 turn | `claude-opus-4-8` | `claude-opus-4-8` ✓ | 3/3 |
  | 9 emails, 3 days, 2 resumes | `claude-opus-4-8` | `claude-opus-4-8` ✓ no drift | 9/9 |

  Two *different* models were served, and neither is the silent default
  (`claude-haiku-4-5`) — chosen deliberately so a persisting M-1 could not hide.
  The 9/9 and 3/3 scores are not a capability signal: these are the first easy emails
  in the plan, and the grader problems (G-*) live further into the run.

- **2026-08-17** — Phase A1. Wrote `docs/benchmark-repair-evidence.md`, the sourced brief
  for the fan-out. Propagated phase 0's findings into the docs that contradicted them:
  `run.sh` setup hint + Python version guard, `BENCHMARK_RESULTS.md` staleness banner and
  corrected setup line, `docs/PROJECT_MAP.md` start-here pointers, `RECAP.md` banner
  correcting its "the machinery is solid" verdict. Moved this register from
  `.claude/plans/` to `docs/` after finding `~/.gitignore_global` excludes
  `.claude/plans/` — it would never have been committed.

- **2026-08-17** — Phase A2. Ran the six-way read-only category fan-out and merged it here.
  **50 findings**: G 10, A 6, O 9, C 9, K 8, V 8. Severity 18 blocks-measurement /
  26 distorts-measurement / 6 slows-work; 49 free-offline to verify, 1 needs-one-live-run.
  Sections kept in full at `docs/_repair/{G,A,O,C,K,V}.md` (4,368 lines); this register
  carries the compressed form and points at them for the working. Nothing outside this file
  was written; no code, corpus, script or other doc was touched; no live run was made.

  **Three cross-category contradictions resolved** (see "Cross-category resolutions"):
  the Haiku-above-Opus anomaly (G explains the band, K explains the confound, A's rank
  inversion is a lever artifact); the opus tool trace (V's completeness result stands,
  O's opus-loss claim withdrawn as a per-day attribution artifact); §4.3 (bounded, not
  closed — majority keyword-discoverability, kind mismatch smallest by an order of
  magnitude, attribution arm unevidenced).

  **The phase table changed.** O moves from phase 1 to phase 6; G becomes the first
  substantive phase; a new **phase 1 "freeze the record"** is inserted ahead of everything;
  V moves ahead of K because V-1's construct decision determines what the corpus pass is
  for. Phase 1.5 splits: the answer-key audit is free and largely delivered by K, the
  hand-grade of model behaviour is not free and moves into phase 7. *The merge argued this
  from two premises; the verifier pass below cut it to one. G-4 carries it.*

  **20 corrections to the evidence brief recorded without editing it**, including six stale
  code anchors, a fourth complete run log (`past/claude-sonnet-4-5.md`, `SCORE 102/176`)
  the brief's §2.1 table omits, and §2.6's unsound inference from the attribution warning
  count. Four operational hazards promoted to a box at the top of this file, the first being
  that the documented "reproduce" path (`scripts/recover_corpus.py`, `scripts/fix_match.py`)
  fetches production and rewrites `corpus/nodes/` in place.

- **2026-08-17** — Phase A, verifier pass. `docs/_repair/VERIFY.md`. ~200 claims checked:
  6 headline numbers re-derived with independent scripts, ~50 further numeric claims,
  ~145 `file:line` anchors opened, 8 grep assertions re-run. **~180 confirmed, 2
  substantively wrong, 8 overstated, 3 unfalsifiable.** All six headline numbers reproduce
  to the digit — G-4's 92/167 and the whole policy table cell-for-cell, G-1's 32/125 and
  8%/48%, A-1's 55/57/43, V-1's 43,474 chars and 18,389-char gap, V-3's 64/167, K-2's
  87-of-127. Merge fidelity clean: all 50 IDs present, zero severity drift, no lost anchors,
  0 of ~145 anchors fabricated.

  **The two substantive errors both fed the phase table, and both are corrected in place.**
  (1) The hazard box called the corpus damage *permanent*. It is not — `corpus/nodes` is 16
  git-tracked files, clean and unchanged since `24331fb`, so `git checkout -- corpus/nodes`
  restores them byte-for-byte. That claim was the sole basis for labelling phase 1
  "time-critical"; it is now "do it early". (2) O-1's "no run already paid for can ever be
  re-graded" is too strong — the register performs three such re-grades itself (G-2, A-1,
  A-6) and the verifier reproduced two from the logs alone. The defensible limit is narrower:
  no re-grade under a rule that would newly admit objects the log never rendered, which still
  blocks G-1, G-3 and G-8. **The phase re-ordering now rests on G-4 alone**, which was
  independently confirmed.

  Also applied: C-1's haiku levers are seven values (`urgency_horizon ∈ 1..7`), not three,
  and the 785 now travels with the grid that defines it; A-1's lever-artifact call upgraded
  from correlational to causal with a model-independent ceiling table (haiku capped nine
  points below opus/sonnet by the serve plan alone, before any model behaviour); resolution
  2's "proves no `get_email` was dropped" softened to a tight lower bound and its
  serialisation inference dropped; G-3's writing-style clause dropped as unmeasurable;
  G-1's prompt claim corrected — `runner.py:107` *does* say "Never leave duplicates", which
  sharpens G-1 (told the right rule, graded by a different one); K-2 retitled off "only
  coherent at `daily_max=21`" since its own table shows 30 strictly better; K-7's "both" →
  "two of four"; six off-by-N anchors; three phase-0 anchors labelled `@24331fb`.

  **New, found by the verifier and owned by no section:** `runner.py:101` instructs the model
  *"Over-acting is the most common mistake; when in doubt, do nothing."* V-3 shows a null
  model scores 38.3%; that line makes claiming the floor the *instructed* strategy. Recorded
  under V-3.

  **Nothing in this pass changed a severity or moved a finding out of `open`.** The
  evidentiary base is sound; what needed editing was prose that over-read it.

  **Not done:** sign-off on the revised phase table. Phase 1 does not begin without it.

- **2026-08-18** — Scope narrowed by decision. V-1 (retrieval span not exercised) and V-5
  (`sb/scale.py` broken) move out of phase 4 and into a **new phase 8**, after the paid
  run. Phases 1-7 are now unambiguously a **repair of the benchmark we have**, not an
  attempt to make it harder. Phase 4 keeps the lever/config pinning and the reporting
  fixes; phase 5's corpus pass is explicitly bounded to making authored questions
  answerable and keywords discoverable.

  Rationale, recorded because it is a deliberate trade: a forecast off the recorded logs
  puts a frontier model near 85-90% after the repair, which suggests the task may be
  trivial once the grader stops hiding capability. That forecast is not a reason to
  redesign now. A measured trivial score is a result, is publishable, and is a far better
  input to a redesign than an estimate is. The alternative — building difficulty in before
  the first honest measurement — risks tuning the benchmark against a number nobody has
  ever seen.

  Supporting measurement (offline, from the existing logs): among failures where the
  grader found more than one matching object, the model had an object on the **exactly
  correct date** in 6/10 (opus), 5/6 (sonnet) and 9/9 (haiku) cases. Date accuracy across
  all work the grader could see: opus 34/48 (71%), sonnet 37/49 (76%), haiku 44/48 (92%).
  Haiku ran the lever the corpus was authored for, which is why its rate is the outlier.

- **2026-08-18** — Provenance audit before phase 1, and a baseline run moved to the front.
  Added **C-1b**: the model label on every surviving artifact is asserted, not observed,
  because M-1 and M-2 were both live when the runs happened. Corpus binding is verified two
  independent ways (167/167 exact email-id resolution per log, plus C-1's unique serve-plan
  replay); model identity cannot be verified from any artifact and no artifact can settle it.

  Phases **1a/1b/1c** inserted: build the minimum re-gradeable capture slice (O-1, O-3, and
  the unmatched-object half of O-5), prove it with a bounded smoke, then run one certified
  baseline at seed 42 / `daily_max=5` for direct comparability with the historical logs.
  O-1, O-3 and O-5 move from phase 6 to 1a; phase 6 keeps the remainder. Phase 7 is no
  longer "the only paid phase".

  This reverses the sign-off's O-timing decision, deliberately and for a stated reason: that
  decision assumed no run would happen before phase 7, which made O's payoff land too late
  to matter. Once a run is happening, building capture first converts one paid run into a
  permanently re-gradeable asset instead of a fourth artifact that cannot be re-scored.
