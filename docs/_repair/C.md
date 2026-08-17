# C — config and provenance

Scope: run levers and corpus identity not being stamped into artifacts; the stale
`BENCHMARK_RESULTS.md`; the empty `past/claude-opus-4-8.md`; and the general problem
that a saved log cannot be trusted to say what produced it.

Phase 0 (register M-2) added a config stamp. What it stamps now, precisely:

| | `sb/live/runner.py` |
|---|---|
| header, printed before day 1 | `:403-408` — requested model, driver, seed, non-empty day count, email count, start date, `daily_min`-`daily_max`, `urgency_horizon`, corpus **path** |
| footer, printed next to the score | `:529-531` — served model, seed, `daily_min`-`daily_max`, email count, corpus **path** |

Stamped nowhere, header or footer: corpus **content** identity (only the directory path
is printed), `urgency_horizon` in the footer, `--start`/`--days`/`--limit`/`--reasoning`/
`--timeout` in the footer, the git revision of `sb/`, the driver CLI's version, the
`SYSTEM_PROMPT` version, a wall-clock timestamp. Neither line is machine-readable, and
neither is present in any of the four saved artifacts, all of which predate phase 0
(they entered the repo at `24331fb`, 2026-07-26; phase 0 is `3956826`, 2026-08-17).

That is the boundary of what phase 0 closed. Everything below is what remains.

---

## C-1 The levers of the only three runs survive by accident, and two documented commands destroy them
Status: open
Severity: blocks-measurement
Cost to verify: free-offline

**What's wrong.** None of the four saved artifacts records its scheduler levers — they
predate the phase-0 stamp — so the config of every run the project owns exists only as
an inference. That config is still recoverable today, because each log prints its own
serve plan and `corpus/` is byte-identical to the commit that added the logs, but the
recovery works only while the corpus stays frozen. Phase 4 is a corpus authoring pass,
and `BENCHMARK_RESULTS.md` §4 documents two commands that rewrite `corpus/nodes/` in
place before anyone reaches phase 4.

**Evidence.**
- The saved headers carry no lever line. `outputs/opus.md:2-5` is the whole header —
  `║ model claude-opus-4-8 via claude` (`:3`) and
  `║ seed 42 · 57 days · 167 emails · start 2026-06-01` (`:4`) between the box rules.
  Same shape in `outputs/sonnet.md:2-5` and `past/claude-haiku-4-5.md:13-16`. The levers
  line at `sb/live/runner.py:406-407` and the footer stamp at `:529-531` were added by
  `3956826`; `git show 3956826 -- sb/live/runner.py` shows both as `+` lines.
- Each log does contain a full serve plan: `runner.py:79` prints
  `[{i}] {eid} · served {served}` per graded email. Extracting those 167 `(email_id,
  serve_date)` pairs and comparing them against plans rebuilt from the current corpus
  recovers the levers uniquely for two of the three runs:

  | artifact | pairs | distinct serve days | plan digest | lever combos matching, out of 785 feasible |
  |---|---|---|---|---|
  | `outputs/opus.md` | 167 | 57 | `832b0b44bce0` | exactly one: `daily_min=1 daily_max=5 urgency_horizon=7` |
  | `outputs/sonnet.md` | 167 | 57 | `832b0b44bce0` | exactly one: `daily_min=1 daily_max=5 urgency_horizon=7` |
  | `past/claude-haiku-4-5.md` | 167 | 16 | `4dd378ec78f1` | three: `daily_min=1 daily_max=21 urgency_horizon ∈ {3,5,7}` |

  (Search space: `daily_min` 1-5 × `daily_max` 1-29 × `urgency_horizon ∈ {3,5,7,10,14,21}`,
  seed 42, start 2026-06-01, `n_days=200`; digest = sha256 of `id@date` pairs in serve
  order, first 12 hex. Script:
  `/private/tmp/claude-501/-Users-jamesoc-dev-SecretaryBench/fe67c85a-9c16-441a-99a9-a4b047c35b38/scratchpad/matchlog.py`.)
- The recovery depends on the corpus being unchanged, and it is:
  `git diff 24331fb HEAD --stat -- corpus/` → empty. `git log --all --oneline -- corpus/`
  shows `24331fb` (2026-07-26) as the only commit touching `corpus/` since
  `4f24e2a` ("reset: clear the corpus to a blank authoring slate", 2026-06-02).
- Re-grading needs the serve plan, not just the score: `runner.py:506` builds each
  email's grading context as `ctx = Context(plan.serve_date[eid], plan.anchors)`, and
  `plan.anchors` comes from `build_plan(..., levers=levers)` at `runner.py:376`.
- The destruction path is documented, not hypothetical. `BENCHMARK_RESULTS.md:146-147`:
  `scripts/recover_corpus.py   # pull+recover -> corpus/nodes/` then
  `scripts/fix_match.py`. `scripts/recover_corpus.py:25,32` fetches
  `https://secretarybench.vercel.app/api/nodes` live at run time;
  `scripts/fix_match.py:21` — `NODE_DIR = Path("corpus/nodes")` — "Reads the recovered
  corpus in corpus/nodes and rewrites it in place" (`:19-20`). Either command run today
  replaces the corpus that the three surviving logs are attributable to.

**Why it matters.** The register's phases 1.5 and 2 both re-use runs already paid for:
hand-grading a sample and re-scoring old runs against a changed grader. Both need to
know which email was served on which day. Right now that mapping is reconstructible;
after any corpus edit it is gone, and the three runs become a set of PASS/FAIL strings
with no recoverable question behind them.

**Options.**
1. Record the recovered provenance into the register now (levers, plan digest, corpus
   commit) as a one-off note, and treat the artifacts as read-only history. Cheap and
   immediate; freezes an inference rather than a measurement, and does nothing for the
   next run.
2. Reconstruct the artifacts once: replay each log against its recovered plan and write
   a machine-readable sidecar (`outputs/opus.provenance.json`) before touching the
   corpus. Preserves re-gradeability across a corpus change; costs a small tool, and
   the haiku sidecar has to record `urgency_horizon` as one of three values.
3. Gate corpus edits: refuse `fix_match.py` / any corpus write while an unrecorded
   artifact exists. Structural, but adds a gate to a repo that currently has none, and
   it constrains the phase-4 authoring pass that the register wants to be free.
4. Accept the loss: declare the three runs superseded, do the corpus pass, and re-run
   the roster in phase 6 with the phase-0 stamp in place. Cleanest story; makes phase 6
   the only source of data and removes the offline baseline phases 1.5 and 2 rely on.

**Overlaps with:** O (no machine-readable output, no store dump), K (phase 4 is the
corpus pass that triggers this), M-2.

**Open questions.**
- Does an off-repo copy of the pre-`24331fb` working tree exist (the four logs were
  produced between the 2026-06-30 smoke and the 2026-07-26 commit by code that was
  never committed at the time)?
- The digest match proves the *scheduling-relevant* structure is identical, and
  `git diff` proves byte-identity for `corpus/`; is that enough to declare the grader
  identical too, given `24331fb` also rewrote `sb/live/runner.py`? See C-6.

---

## C-2 `urgency_horizon` is absent from the score-line stamp and moves 89% of serve dates
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** The footer stamp — the one line a human copies next to a score —
records `daily_min`-`daily_max` but not `urgency_horizon`, and every hand-written
provenance header in `past/` omits it too. Measured on the current corpus,
`urgency_horizon` is the *largest* single determinant of the serve plan: moving it from
7 to 14 changes the serve date of 148 of 167 emails. The non-empty day count that
everyone reads off the header is nearly blind to it, so the omission is invisible.

**Evidence.**
- `sb/live/runner.py:529-531`, the whole stamp:
  ```
  # Stamp the config next to the score so a saved log can't be mislabelled later.
  print(f"  {DIM}served {resolved_model or model} · seed {seed} · daily "
        f"{levers.daily_min}-{levers.daily_max} · {len(order)} emails · {corpus_dir}{RESET}")
  ```
  `urgency_horizon` appears at `:406-407` (header) and nowhere in the footer.
- The hand-written headers omit it as well. `past/claude-haiku-4-5.md:4`:
  `- **Corpus:** corpus/ (167 emails, sha d737d44e14dc7d20) · seed 42 · days 30 · daily_max 21`.
  `past/claude-sonnet-4-5.md:4`:
  `- **Corpus:** 176 emails (sha 809d389794dd79a9) · seed 42 · days 30 · daily_max 21`.
- Measured effect at `daily_min=1, daily_max=5`, seed 42, `n_days=200`:

  | comparison | emails whose serve date changes | max shift |
  |---|---|---|
  | `urgency_horizon` 7 → 14 | **148 / 167** | 34 days |
  | `urgency_horizon` 5 → 7 | 119 / 167 | — |

  All three of `{5, 7, 14}` still produce exactly 57 non-empty serve days, so the header's
  day count cannot distinguish them. Plan digests differ: `832b0b44bce0` (hor 7) vs
  `4c01f769e0e3` (hor 14).
- By contrast `--days` is inert once feasible: `daily_max=5` at `n_days` 60 / 90 / 200 all
  give digest `832b0b44bce0`. It only gates feasibility — minimum feasible `--days` is 57
  at `daily_max=5`, 68 at 4, 16 at 21, and `daily_max=3` is infeasible at any ceiling.
  Reproduce: `/private/tmp/claude-501/-Users-jamesoc-dev-SecretaryBench/fe67c85a-9c16-441a-99a9-a4b047c35b38/scratchpad/plans2.py`.
- The framing that plausibly caused the omission: `docs/design/SERVING_AND_SCHEMA.md:139`
  — `| `urgency_horizon` | Buffer before `date` deadlines; mostly a safety knob,
  secondarily affects clustering. |`

**Why it matters.** Answer keys resolve against the serve date (`runner.py:506`), so a
lever that moves 89% of serve dates changes what is being asked, not merely when it is
asked. Two runs stamped identically — same model, seed, `daily 1-5`, 167 emails, same
corpus path — can have been asked different questions. That is exactly the failure the
stamp was introduced to prevent.

**Options.**
1. Add `urgency_horizon` (and `--start`) to the footer string. One-line change; still
   prose in a human log, still not machine-readable, still trusts the printer.
2. Print a single derived plan digest (sha of `id@date` in serve order) in the footer.
   Identifies corpus *and* all levers at once and is checkable offline; a digest tells
   you two runs differ without telling you how, and it needs the corpus to interpret.
3. Emit the full argparse namespace as a JSON block. Complete and machine-readable;
   duplicates the work O is scoping for run artifacts, so the two should land together
   or one will be thrown away.
4. Remove the lever knobs from the CLI and pin one config in code. Kills the class of
   error outright; also kills `sb.scale` span experiments and the phase-6 flexibility.

**Overlaps with:** C-1, C-4, O, V.

**Open questions.**
- `urgency_horizon` is the one lever whose effect the design doc understates — is the
  scheduler's use of it (`sb/scheduler.py`, deadline buffer) doing what the doc says?
  That is a scheduler question, not a config one, but it decides whether the doc line
  or the code is the thing to fix.

---

## C-3 Corpus satisfiability is lever-dependent, and the oracle gate is hardcoded to one lever setting
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** The register and the brief both state flatly that every answer key is
satisfiable, on the strength of `sb.scale` printing `oracle: 167/167 = 100%`. That gate
runs at the argparse defaults only — `sb/scale.py:111` calls `build_plan` with no
`levers` argument and `sb.scale` exposes no lever flags. Sweeping the lever space, 12 of
93 feasible settings leave an answer key that the *perfect* model cannot satisfy.

**Evidence.**
- `sb/scale.py:111`:
  `plan = build_plan(corpus, start_date=date(2026, 6, 1), seed=args.seed, n_days=args.days)`
  — no `levers=`, so `Levers()` defaults `(1, 5, 7)`. `sb.scale --help` offers only
  `--filler`, `--seed`, `--days`, `--dst`.
- Sweep over `daily_min` 1-5 × `daily_max ∈ {5,8,12,21}` × `urgency_horizon ∈ {3,5,7,10,14}`,
  seed 42, `n_days=200`: **12 of 93 feasible settings score 166/167**, always failing the
  same email, `Marketing-campaign-new-product-delay.serena-williams-reschedule`.
  Failing settings include `(1, 8, 10)`, `(2, 5, 7)`, `(4, 12, 7)`, `(5, 8, 7)`.
- The grader's own reason at `(daily_min=2, daily_max=5, urgency_horizon=7)`:
  ```
  {'passed': False, 'label': 'move ~marketing',
   'expected': 'event ~"marketing" @ Wed Sep 30',
   'actual': '"LeBron James marketing campaign scheduled" Wed Sep 30 9 AM; "Giano Ronaldo marketing campaign " Sun Aug 16 9 AM',
   'reason': 'found 2 matching, expected exactly 1 (duplicate / double-booked)'}
  ```
  The op is `{"move": "Serena William marketing campaign ", "match": ["marketing"],
  "on": {"eq": "@SW_Marketing_Campaign+1w+4d"}}` in
  `corpus/nodes/Marketing-campaign-new-product-delay.json`. Three obligations in that node
  contain "marketing"; whether all three are in the cumulative pool by that day depends on
  the serve order, which depends on the levers.
- A second, unrelated hardcode: `scripts/fix_match.py:114-115` gates the corpus at
  `n_days=730, Levers(daily_max=21)` — a third lever setting, matching neither the
  argparse defaults nor `sb.scale`.
- Reproduce (about 4 minutes offline):
  `/private/tmp/claude-501/-Users-jamesoc-dev-SecretaryBench/fe67c85a-9c16-441a-99a9-a4b047c35b38/scratchpad/` — the sweep is inlined in the session transcript; the shape is
  `build_plan(corpus, start_date=date(2026,6,1), seed=42, n_days=200, levers=Levers(dmin,dmax,hor))`
  then `run(corpus, plan, oracle_model, store=Store(corpus))`.

**Why it matters.** "The corpus is oracle-clean" is the project's standing proof that a
score is a capability measurement rather than an artifact of an impossible key. That
proof is currently conditional on a lever setting that is not part of the claim, not
stamped in any artifact, and freely changeable from the CLI. A phase-6 comparison run at
`--daily-min 2` would be measuring a corpus with a provably unwinnable email, and nothing
in the workflow would say so.

**Options.**
1. Add lever flags to `sb.scale` so the gate can be run at the levers a live run will
   use. Small; makes the gate correct only if someone remembers to match them.
2. Make the gate sweep: `sb.scale` runs the oracle across a lever grid and reports the
   worst case. Catches the whole class; slower, and turns a green/red gate into a
   distribution someone has to interpret.
3. Pin the levers as constants and delete the CLI flags (see C-2 option 4), making the
   single gate run sufficient by construction. Removes the discrepancy; removes the
   difficulty dial the design doc (`SERVING_AND_SCHEMA.md:132-140`) treats as a feature.
4. Treat it as a corpus defect only — fix the colliding `match` keyword in that node and
   re-check. Fixes today's instance; leaves the gate blind to the next one.

**Overlaps with:** G (cumulative-pool keyword collisions, lint #5's blind spot),
K (`match` keyword authoring), C-2.

**Open questions.**
- Is `(1, 5, 7)` special, or merely lucky? The default is clean, but so are 81 of the 93
  settings — nothing suggests the default was chosen because it is clean.
- Does the same lever-dependence apply to feasibility of `by:` deadlines elsewhere, or is
  `marketing` the only collision that the serve order can expose?

---

## C-4 `sb.analyze` takes the levers by hand and silently reports different spans if they are wrong
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** `sb.analyze` rebuilds the serve plan from its own CLI arguments and
never reads the levers out of the log it is analysing — not even now that the log prints
them. Its defaults disagree with the runner's, so the obvious invocation is wrong for
most logs. It does not warn: it prints a complete, confident span-vs-accuracy report
computed against a plan the run never used.

**Evidence.**
- `sb/analyze.py:76-83`: `ap.add_argument("log")`, `--corpus` default `build/scaled`,
  `--seed` 42, `--days` 300, `--daily-min` 1, `--daily-max` 5, `--urgency-horizon` 7.
  Runner defaults are `--corpus corpus`, `--days 60` (`runner.py:544,542`). `parse_log`
  (`analyze.py:34-72`) extracts only PASS/FAIL and `searched`; nothing reads the header.
- Same log, two lever settings, no warning either time:
  ```
  $ .venv/bin/python -m sb.analyze past/claude-haiku-4-5.md --corpus corpus
  untagged    42% (19)    80% (5)        ·
  0-50           19      42%        0%
  50-100          5      80%        0%
  overall needle accuracy: 50% over 24 needles

  $ .venv/bin/python -m sb.analyze past/claude-haiku-4-5.md --corpus corpus --daily-max 21 --days 30
  untagged    31% (13)    70% (10)    100% (1)
  0-50           13      31%        0%
  50-100         10      70%        0%
  100+            1     100%        0%
  overall needle accuracy: 50% over 24 needles
  ```
  The second is the correct config for that log (C-1 recovers `daily_max=21`). The span
  distribution moves by 6 needles between bins and the `100+` bin appears only in the
  correct run; the overall figure is span-independent and hides the error.
- `BENCHMARK_RESULTS.md:159-160` already knows this is a trap — `# optional richer report
  (tier × span) — MUST pass the same levers:` — which makes it a documented manual
  discipline rather than a checked one.
- Both invocations report every needle as `untagged`, corroborating brief §2.9: the
  schema field is `tier` (`sb/schema.py:81`, `T1|T2|T3`) while `analyze.py:104` reads
  `entry.get("reasoning_tier", "untagged")` from the needle registry, and
  `analyze.py:31` still carries `TIER_ORDER = ["T1", "T2", "T3", "T4"]`.

**Why it matters.** The span-vs-accuracy grid is the benchmark's headline research claim
(`analyze.py:2-4`: "does a needle's correctness fall as the fact it needs sits further
back"). A tool that computes that grid against the wrong serve plan and says nothing
produces a plausible answer to a question nobody asked. `RUN_RESULTS.md`'s published
tier × span grid was produced by this tool and its levers were never recorded.

**Options.**
1. Have `sb.analyze` parse the levers from the runner's header/footer line and refuse to
   run if the operator's flags disagree. Uses the phase-0 stamp; only works for logs
   produced after phase 0, i.e. none of the four that exist.
2. Have the runner emit a machine-readable run record and have `sb.analyze` consume that
   instead of a human log. Removes the re-parse entirely; it is the same artifact O is
   scoping, so it should be one decision, not two.
3. Make the lever flags required (no defaults) so a wrong invocation is an error rather
   than a plausible report. Trivial; converts a silent wrong answer into an annoying
   prompt, and still trusts whatever the operator types.
4. Have `sb.analyze` verify its rebuilt plan against the `· served <date>` lines the log
   already prints, and abort on mismatch. Self-checking against data already in every
   log including the four old ones; needs the log format to stay stable.

**Overlaps with:** O (log re-parsing, no machine-readable output), V (tier data unread,
span reporting), C-2.

**Open questions.**
- `RUN_RESULTS.md`'s grid: was it produced at levers that matched its run? The run
  predates the levers CLI entirely (`Levers` enters at `24331fb`), so probably yes by
  default — but the log is gone (C-8), so it cannot be checked.

---

## C-5 Corpus identity is asserted by a hash whose algorithm exists nowhere in the repo
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** Three places in the durable record identify the corpus by a truncated
hash. No code in the repository computes a corpus hash, no doc states the algorithm, and
six plausible reconstructions fail to reproduce the value recorded for the corpus
currently on disk. The one field intended to pin corpus identity cannot be checked
against anything.

**Evidence.**
- Every occurrence of a corpus hash in the repo:
  ```
  BENCHMARK_RESULTS.md:81   | corpus sha256 (node files) | `809d389794dd79a9…` (post match-keyword fix) |
  BENCHMARK_RESULTS.md:205    Final corpus sha256 `809d389794dd79a9…`.
  past/claude-sonnet-4-5.md:4 - **Corpus:** 176 emails (sha 809d389794dd79a9) · seed 42 · days 30 · daily_max 21
  past/claude-haiku-4-5.md:4  - **Corpus:** corpus/ (167 emails, sha d737d44e14dc7d20) · seed 42 · days 30 · daily_max 21
  ```
  (`grep -rn "sha\b\|sha256\|d737d44e\|809d3897" --include="*.md" --include="*.py" --include="*.sh" .`)
- No implementation: `grep -rn "sha256\|hexdigest\|blake2" scripts/*.py sb/*.py sb/live/*.py`
  returns nothing.
- Six candidate algorithms over the 15 files in `corpus/nodes/`, none matching
  `d737d44e14dc7d20` (the value recorded for a 167-email corpus):

  | algorithm | first 16 hex |
  |---|---|
  | concat raw bytes, name-sorted | `0adbe090d13dde1a` |
  | concat name + bytes | `03e0d963b9866d8f` |
  | sha of per-file shas | `89002c2cd4d05b02` |
  | canonical JSON `{name: obj}` | `de7db3834e377913` |
  | canonical JSON list | `d0a8dc20579fa1a7` |
  | concat decoded text | `0adbe090d13dde1a` |

  Script: `/private/tmp/claude-501/-Users-jamesoc-dev-SecretaryBench/fe67c85a-9c16-441a-99a9-a4b047c35b38/scratchpad/tryhash.py`.
- What the runner stamps instead is a path, not content: `runner.py:531` prints
  `{corpus_dir}` — the string `corpus`. Two runs against completely different corpora at
  the same path stamp identically.
- The corpus *is* verifiable by other means — `git diff 24331fb HEAD --stat -- corpus/`
  is empty — so this is a broken mechanism, not a lost corpus.

**Why it matters.** `daily_max`, `seed` and the model can all be recovered or inferred;
corpus content is the one thing that cannot be reconstructed from a log once the files
change. The record's only defence against that is a number nobody can recompute. It also
makes the two `past/` headers unfalsifiable: they *look* like verified provenance and are
in fact hand-typed prose.

**Options.**
1. Define and implement one canonical corpus digest (`sb.corpus_hash`), print it in the
   runner header/footer, and note that the two historical values are unverifiable. Cheap
   and durable; the old numbers stay dead, so `past/*.md` gains a caveat, not a check.
2. Drop content hashing and stamp the git revision of `corpus/` instead
   (`git rev-parse HEAD:corpus`). Free, exact, already maintained; useless for the
   uncommitted or generated corpora (`build/scaled*`) that `sb.scale` produces.
3. Stamp both: git tree hash when clean, content digest when dirty or generated. Covers
   every case; two mechanisms to keep honest, and a "dirty" stamp still does not say
   *what* was dirty.
4. Stamp the derived plan digest instead of the corpus (C-2 option 2). One number covers
   corpus and levers together; it does not detect corpus edits that leave scheduling
   unchanged — including every `match`-keyword change, which is exactly what phase 4 and
   `fix_match.py` do.

**Overlaps with:** C-1, C-2, O.

**Open questions.**
- Does `d737d44e14dc7d20` correspond to the corpus in `24331fb` under some algorithm not
  tried here (e.g. hashing the webapp export JSON rather than the node files, which is
  what `recover_corpus.py` fetches)? If so the value is recoverable rather than dead.

---

## C-6 The whole evidentiary base entered the repo in one commit titled "." that also rewrote the harness
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** `24331fb` added the durable record, the entire corpus, all five run
artifacts, both corpus-mutating scripts and a 190-line rewrite of the runner, under the
commit message ".". The four logs were therefore produced by working-tree code that was
never committed while it was in use; only its end state is in git. Nothing in history
attributes any score to a harness revision.

**Evidence.**
- `git show 24331fb --stat` (author Tyrost, 2026-07-26, message `.`), 26 files, +9257/-25:
  ```
  BENCHMARK_RESULTS.md          |  193 ++++
  corpus/nodes/*.json           |  (15 files, ~4400 lines)
  outputs/opus.md               | 1144 ++++
  outputs/sonnet.md             | 1143 ++++
  past/claude-haiku-4-5.md      |  991 ++++
  past/claude-opus-4-8.md       |    0
  past/claude-sonnet-4-5.md     | 1047 ++++
  sb/analyze.py                 |   10 +-
  sb/live/runner.py             |  190 +++-
  scripts/fix_match.py          |  136 +++
  scripts/recover_corpus.py     |  137 +++
  ```
- The runner change in that same commit is behavioural, not cosmetic:
  `git show 24331fb -- sb/live/runner.py` adds the `codex` driver and `infer_driver`,
  adds `Levers` to the scheduler import, adds `_limit_reset_wait` (usage-window pausing),
  and rewrites `build_cmd` including the comment `NO --tools "" (in current CLI that
  means "zero tools available", which blocked the MCP tools)` — the fix
  `BENCHMARK_RESULTS.md:26-31` dates to the 2026-06-30 smoke, i.e. four weeks before the
  commit.
- `CLAUDE.md` (project instructions): "**Never change the grader and the corpus in the
  same commit.** If both move, a score change cannot be attributed to either." — and
  `CLAUDE.md` is itself gitignored (`.gitignore:30`), so that rule has never been visible
  to the author who broke it.
- The 176-email corpus was never committed at all: `git log --all --oneline -- corpus/`
  goes `4f24e2a` (blank slate, 2026-06-02) → `24331fb` (15 files, 2026-07-26), and
  `git ls-tree -r --name-only 4f24e2a -- corpus/` returns only `corpus/nodes/.gitkeep`.
  So `past/claude-sonnet-4-5.md`'s 102/176 run is attributable to no corpus in this repo.

**Why it matters.** Register phase 2 will change the grader and re-score. The only
baseline it can compare against is a set of logs whose harness revision is unknown and
whose corpus, for one of them, never existed here. A score delta measured against that
baseline cannot be attributed to the grader change.

**Options.**
1. Reconstruct what can be reconstructed and write it down: pin `24331fb` as the harness
   revision for the three 167-email logs (best available), and mark
   `past/claude-sonnet-4-5.md` as un-attributable. Honest and free; "best available" is
   an assumption, since the code moved during the runs.
2. Retire the pre-phase-0 artifacts as baselines entirely; use them only as text for the
   phase-1.5 hand-grade, where the serve plan (C-1) is all that is needed. Removes the
   attribution problem; costs the only comparison points that exist.
3. Add a commit-time guard (pre-commit hook or CI check) that refuses a commit touching
   both `corpus/` and `sb/`. Prevents recurrence; does nothing for the existing history,
   and `CLAUDE.md` being gitignored means the rule still is not visible to contributors.
4. Move the operating rules out of the gitignored `CLAUDE.md` into a tracked
   `CONTRIBUTING.md`. Makes the rule reach the team; a rule that is only written down is
   the thing that already failed once.

**Overlaps with:** C-1, C-8, M-5.

**Open questions.**
- Did `SYSTEM_PROMPT` change between the four runs? It is unchanged *by* `24331fb`
  (`git show 24331fb -- sb/live/runner.py` contains none of its lines), but the runs
  span 2026-06-30 to 2026-07-26 of uncommitted work, so "unchanged in the commit" does
  not mean "unchanged across the runs".

---

## C-7 The stamp prints "served <model>" from the requested model when nothing was ever served
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** The footer's served-model field falls back to the *requested* model
whenever no turn produced a resolved model, and prints it under the word "served" with no
qualifier. Separately, a mid-run model drift overwrites the recorded value, so the footer
names only the last model of a run that used two. Both cases produce a stamp that asserts
more than the run established — the exact failure M-2 was introduced to close.

**Evidence.**
- `sb/live/runner.py:530`: `print(f"  {DIM}served {resolved_model or model} · seed {seed} · daily "`.
  `resolved_model` is initialised `None` at `:410` and set only inside the success branch
  at `:456-458` (`if rmodel:`). If every day errors out (`:491-496` prints
  `ERROR after retries` and `continue`s), or if a driver's stream stops carrying a model
  on the init event, `resolved_model` stays `None` and the footer prints the requested
  model as though it were observed.
- Drift collapses to the last value: `:465-469` prints the `MODEL DRIFT` warning and then
  `resolved_model = rmodel`. The warning is one line in the middle of a 1100-line log; the
  footer stamp at the end contradicts it silently by naming a single model.
- The mismatch test is `rmodel == model or rmodel.startswith(model)` (`:459`). The
  `startswith` is deliberate (register M-3: `claude-sonnet-4-5` → `claude-sonnet-4-5-20250929`)
  but is prefix-loose in both useful and unhelpful directions — a request for a family
  prefix would pass against any member of it.
- No saved artifact exercises this: none of the four logs contains the stamp at all (all
  predate `3956826`).

**Why it matters.** The stamp's whole purpose is that "a saved log can't be mislabelled
later" (`runner.py:529`, verbatim). A field that silently substitutes the requested value
for the observed one restores the original failure mode in the one place a reader is most
likely to trust. An all-errored run is not hypothetical — `BENCHMARK_RESULTS.md:173`
records a run that lost 35 emails to a usage window.

**Options.**
1. Print `served <model>` only when observed; otherwise print `served: not observed
   (requested <model>)`. Minimal and unambiguous; adds a second format the eventual
   machine-readable record has to represent.
2. Record the full set of resolved models and print all of them (`served
   claude-opus-4-8 (days 1-12), claude-sonnet-4-5 (days 13-57)`). Makes drift legible at
   the score line; more state to carry, and per-day model attribution only matters if
   drift is real, which M-3 found it is not on CLI 2.1.233.
3. Make an unresolved model or any drift a hard failure that aborts the run. Strongest
   guarantee; throws away partial runs that are currently salvageable, and a paid run is
   expensive to throw away.
4. Leave the print and rely on the inline warnings, treating the footer as a summary.
   Zero work; the footer is the line that gets copied into `past/*.md` headers, and the
   inline warnings are the lines that do not.

**Overlaps with:** M-2, M-3, O.

**Open questions.**
- What does `--driver codex` put on its init event? `_parse_codex` was unit-checked on
  synthetic streams (register M-3) but no codex run exists here — `codex` is not installed
  on this machine (register M-4) — so the resolved-model path is unexercised for half the
  roster.

---

## C-8 Run logs default to a gitignored path, so artifacts are preserved only by hand
Status: open
Severity: slows-work
Cost to verify: free-offline

**What's wrong.** Every documented run command redirects output into `build/`, which is
gitignored, so the primary artifact of a paid run is created where git will never see it.
Preservation is a manual copy into `outputs/` or `past/` under two different, undocumented
conventions. Three of the record's known gaps — the 0-byte opus file, the vanished 84/176
haiku run, and the ~51% run of brief §4.4 — are all instances of that one mechanism.

**Evidence.**
- `BENCHMARK_RESULTS.md:152-160`: `> build/run_haiku.log 2>&1`, `> build/run_gpt55.log 2>&1`,
  and `sb.analyze build/run_haiku.log`. `RUNNING.md:78` and `RUN_RESULTS.md:12` do the same.
  `.gitignore:9` and `:39` both list `build/`.
- The referenced logs are gone: `past/claude-haiku-4-5.md:6` cites
  `build/run_claude-haiku-4-5_v2.log (completed 2026-07-04)`; `find build -type f` returns
  only `build/scaled0/nodes/*.json`, and `find . -name "*.log" -not -path "./.venv/*"`
  returns nothing.
- `past/claude-opus-4-8.md` was born empty — `git show 24331fb --stat` lists it as `0`
  lines, and it has never been non-empty in any commit.
- The 84/176 haiku artifact was overwritten rather than lost in transit.
  `BENCHMARK_RESULTS.md:173` links `outputs/claude-haiku-4-5.md` for 84/176; the file at
  the analogous path today is `past/claude-haiku-4-5.md`, whose own header says
  (`:5-6`): `- **Score:** `SCORE 98/167 (59%)`` … `(clean retry — supersedes the earlier
  rate-limited 84/176 run on the old 176-email corpus)`. A model-id filename gets reused
  by the next run of that model.
- Two conventions in one commit: `past/<model-id>.md` carries a hand-written provenance
  header plus fenced stdout; `outputs/<nickname>.md` is raw stdout with no header at all.
  The newer, currently-cited artifacts (`outputs/opus.md`, `outputs/sonnet.md`) therefore
  carry *less* provenance than the older ones.
- Broken links are exactly the two the brief names, and no others anywhere in the doc
  surface. Every markdown link target in every tracked `.md` was resolved; the only
  failures are `BENCHMARK_RESULTS.md → outputs/claude-haiku-4-5.md` and
  `→ outputs/claude-sonnet-4-5.md`. The second artifact's *content* does survive, at
  `past/claude-sonnet-4-5.md:5` — `SCORE 102/176 (58%)` — only the link is wrong.
- The ~51% run is not recoverable from this repository. Across every commit on every ref
  including `origin/backups`, exactly four files have ever contained a `SCORE` line:
  ```
  git rev-list --all | while read c; do git grep -l "SCORE [0-9]" $c; done | sort -u
  → outputs/opus.md, outputs/sonnet.md, past/claude-haiku-4-5.md, past/claude-sonnet-4-5.md
    (in each of 24331fb, 3956826, 67b3005)
  ```
  `git fsck --lost-found` reports no dangling objects, `git stash list` is empty, and no
  `.log` file exists on disk.

**Why it matters.** A live run is the only expensive thing this project does
(`CLAUDE.md`: "A full run is ~57 day-turns"). A workflow whose default destination is
gitignored, whose filenames collide across runs of the same model, and whose preservation
step is a human remembering to copy, will keep losing paid runs. It already has, at least
three times.

**Options.**
1. Write run output to a tracked, timestamped path by default
   (`outputs/<model>-<date>-<planhash>.md`) instead of relying on shell redirection.
   Removes the collision and the manual step; adds committed artifacts to the repo at
   ~65KB per run.
2. Keep `build/` as the working destination and add an explicit `--out` flag plus a
   documented archive step. Smaller change, keeps the repo lean; still a manual step,
   which is the thing that fails.
3. Make the runner refuse to start if it cannot write a durable artifact. Guarantees no
   unrecorded paid run; blocks quick smokes unless there is an opt-out, and an opt-out is
   how this happens.
4. Adopt one artifact convention and normalise the existing five (header + fenced stdout,
   model-id filename with a run discriminator). Makes the record readable; touches saved
   artifacts, which is a history edit and needs its own justification.

**Overlaps with:** C-1, C-6, O.

**Open questions.**
- Do the `build/run_*.log` files still exist outside the repo on this machine (older
  working copy, Time Machine, a `~/Downloads` copy)? That is the only remaining place the
  ~51% run could be.
- `past/claude-opus-4-8.md` has never held content — was it a placeholder for the run
  that later became `outputs/opus.md`, or for a fourth, separate opus run?

---

## C-9 Residual stale claims in the durable record that the phase-A1 banner does not cover
Status: open
Severity: slows-work
Cost to verify: free-offline

**What's wrong.** The phase-A1 banner on `BENCHMARK_RESULTS.md` warns off the §1 config
and §5 results, but several load-bearing claims outside that scope are still wrong, and
`RUN_RESULTS.md` publishes an 86% headline with no banner at all and a reproduce block
that cannot run. The stale content is not confined to numbers: it includes dead CLI flags,
a retired harness described as current, and an unqualified reproducibility claim.

**Evidence.**
- The §1 justification for the pinned config is false for the corpus in the repo.
  `BENCHMARK_RESULTS.md:89-91`: "`daily_max=21` is the **minimum feasible** for the raw
  authored corpus … at the pilot default (`daily_max=5`) the schedule is infeasible."
  Measured on the current 167-email corpus: `daily_max=5` is feasible at `--days ≥ 57`
  and `daily_max=4` at `--days ≥ 68`; only `daily_max=3` is infeasible at any ceiling.
  The banner flags the config as stale but not this claim, and the claim is what a reader
  would use to pick levers for phase 6.
- `BENCHMARK_RESULTS.md:18-19`, above the banner's reach in reading order:
  "> Status: **SMOKE VERIFIED — ready for the full run.**" — contradicted by the banner
  eleven lines above it.
- `BENCHMARK_RESULTS.md:175` still shows `claude-opus-4-8 | — | — | — | running` while
  `outputs/opus.md` holds that run's 90/167 result.
- The §4 "reproduce" block is not reproducible: `scripts/recover_corpus.py:25,32` fetches
  `https://secretarybench.vercel.app/api/nodes` live, and `CLAUDE.md` lists the production
  webapp and the authored corpus in production as do-not-touch. The command therefore
  cannot be run to reproduce a past corpus, only to overwrite the present one (see C-1).
- Dead CLI flag in two docs: `RUNNING.md:75` and `:91`, and `RUN_RESULTS.md:11`, all pass
  `--needles`. `.venv/bin/python -m sb.scale --help` accepts only
  `[--filler FILLER] [--seed SEED] [--days DAYS] [--dst DST]`.
- `RUNNING.md:42` documents a retired harness — `sb.live.runner (hand the model one email
  at a time)` — against `runner.py:412` (`for day_no, batch in enumerate(days, 1)`) and
  `RECAP.md:71-93`, which describe one turn per day. `RUNNING.md:3` states "Everything is
  reproducible. Same seed in, same result out.", which is true of the serve plan and false
  of a live run. `72796af` (2026-06-29) is titled "fix stale run commands in RUNNING.md",
  so this file has already been corrected once and is stale again.
- `RUN_RESULTS.md` has no staleness banner. Its corpus (`:8`, "267 emails = the
  handwritten `corpus/` nodes + 200 … filler + 24 planted tests") sits on a corpus wiped
  by `4f24e2a` (2026-06-02); its tier vocabulary is T1–T4 (`:18`) against
  `sb/schema.py:81` (`T1|T2|T3`); its closing status (`:56`) says the tooling "is **not
  committed yet**". `docs/PROJECT_MAP.md:22` describes it only as "historical Haiku
  retrieval-span pilot results", which does not warn a reader off its 86% headline.
- `run.sh:6` — `#   ./run.sh --model claude-sonnet-4-6 --seed 7` — names a model id that
  appears nowhere else in the repo (`grep -rn "claude-sonnet-4-6"` returns that line only)
  and is not on the `BENCHMARK_RESULTS.md:102` roster that register M-4 verified.

**Why it matters.** The register's phase 5 rewrites `BENCHMARK_RESULTS.md`. Everything
listed here is a claim someone could act on *before* phase 5 — pick `daily_max=21`
because §1 says 5 is infeasible, run `recover_corpus.py` because §4 calls it the
reproduce path, cite RUN_RESULTS.md's 86% as a span result. Each of those is a wrong move
that the current banners do not prevent.

**Options.**
1. Extend the existing banner treatment to `RUN_RESULTS.md` and `RUNNING.md` and correct
   the specific false claims in place. Consistent with what phase A1 already did; more
   banners on more files, and banners are what readers skip.
2. Move superseded documents into `docs/history/` (which already exists for exactly this)
   and leave a stub. Unambiguous about status; breaks `docs/PROJECT_MAP.md` references
   and any external link.
3. Fix only the mechanically checkable items now (dead `--needles` flags, the two broken
   links, the `running` row, the `claude-sonnet-4-6` example) and defer prose to phase 5.
   Small, safe, verifiable; leaves the §1 feasibility claim — the most actionable wrong
   statement here — standing longest.
4. Add a docs test that fails when a documented command's flags do not exist and when a
   markdown link 404s. Catches the whole class going forward; does not touch prose
   claims, which is where the false statements are.

**Overlaps with:** C-1, C-5, V (RUN_RESULTS.md's span claims), M-4.

**Open questions.**
- Is `claude-sonnet-4-6` a real id that the roster is missing, or a typo in `run.sh`?
  Deciding needs a model-availability check, which is a live call.
- `RUN_RESULTS.md:44` attributes an earlier 21% score to a keyword double-count fixed by
  "matching the action word too" — the same class of failure as G's cumulative-pool
  collisions. Is that fix still present in the grader, or was it lost across the DAG
  transition?

---

## Notes on the brief

**§2 claims reproduced.**
- §2.1 scores and spans: `grep -n "SCORE" outputs/opus.md outputs/sonnet.md past/claude-haiku-4-5.md`
  → `90/167 (54%)`, `91/167 (54%)`, `98/167 (59%)`; headers give 57 / 57 / 16 days.
- §2.1 `past/claude-opus-4-8.md` is 0 bytes, and was created empty (`git show 24331fb --stat`).
- §2.1 the two §5 evidence links are dead — and they are the *only* broken markdown links
  anywhere in the tracked doc surface (every link target in every tracked `.md` was resolved).
- §2.10 corpus state and `oracle: 167/167 = 100%` at default levers.
- §2.10 "the 57-day span … confirm[s] those ran at default levers": reproduced and
  strengthened — see below.

**§2 claims I would amend.**
1. **§2.5 "The oracle cannot detect this class of bug" is overstated.** The oracle cannot
   detect keywords a real model would never produce, which is the point §2.5 is making.
   It *does* detect keyword collisions *between obligations*, and does so here: at
   `(daily_min=2, daily_max=5, urgency_horizon=7)` the oracle scores 166/167 with
   `reason: 'found 2 matching, expected exactly 1 (duplicate / double-booked)'` on
   `Marketing-campaign-new-product-delay.serena-williams-reschedule`. The distinction
   matters for G: the oracle is a partial detector for the cumulative-pool problem, not a
   blind one, and lint #5 let this case through (see C-3).
2. **§2.10 and register:194 "every answer key is satisfiable" is true only at the levers
   tested.** 12 of 93 feasible lever settings score 166/167 (C-3). The claim needs the
   qualifier "at `daily_min=1 daily_max=5 urgency_horizon=7`".
3. **§2.1's levers column is under-specified for haiku and over-trusting for its source.**
   `daily_max=21` is independently confirmable from the haiku log's serve plan, but
   `urgency_horizon` is not: three values (3, 5, 7) reproduce that log exactly. The
   brief's only stated source for the haiku levers is a hand-typed header line
   (`past/claude-haiku-4-5.md:4`), which C-5 shows is unverifiable prose in its corpus-sha
   field.
4. **§2.1's "two evidence links point at files that do not exist" is right but incomplete.**
   One of the two artifacts survives: `past/claude-sonnet-4-5.md:5` is the 102/176 sonnet
   run the §5 row cites. The haiku 84/176 artifact is genuinely gone — the file at that
   name was overwritten by a later 98/167 run (`past/claude-haiku-4-5.md:6`).

**§4 items I believe are now established.**
- **§4.4, the ~51% run: ruled out as recoverable from this repository.** Across every
  commit on every ref (including `origin/backups`), exactly four files have ever contained
  a `SCORE` line, and none of them scores 51%:
  ```
  git rev-list --all | while read c; do git grep -l "SCORE [0-9]" $c; done | sort -u
  ```
  `git fsck --lost-found` shows no dangling objects, `git stash list` is empty, and no
  `.log` exists on disk. Git archaeology is exhausted; only an off-repo copy could exist.
  C-8 identifies the mechanism (`build/` is gitignored and every documented run command
  redirects into it), which also explains why no such artifact would have been committed.
- **New, and not in §4: the levers of the surviving runs are recoverable offline, for now.**
  `outputs/opus.md` and `outputs/sonnet.md` match exactly one of 785 feasible lever
  combinations — `daily_min=1 daily_max=5 urgency_horizon=7`, the full argparse defaults —
  by replaying the `· served <date>` lines the logs already print against rebuilt plans
  (C-1). This upgrades §2.10's day-count inference to a unique identification and, as a
  side effect, confirms the current `corpus/` is the corpus those runs were served from.
  It stops working the moment the corpus changes.
