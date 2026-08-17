# O — run artifacts and observability

Category O of the repair register. Scope: what a run leaves behind, and whether it is
enough to re-derive a score without paying for the model again.

All measurements below were taken on 2026-08-17 against the working tree at
`debug/model-resolution` (`67b3005`). Throwaway analysis scripts live outside the repo;
each finding names a command that reproduces its numbers from committed files only.

---

## O-1 The harness writes nothing; the surviving artifacts are hand-copied stdout
Status: open
Severity: blocks-measurement
Cost to verify: free-offline

**What's wrong.** `sb.live.runner` has no file output of any kind — no `--out`, no
`--json`, no `open(..., "w")` anywhere in the module — so every run's record is whatever
a human happened to redirect and paste into a markdown file. The store process that
holds the entire run's state is `terminate()`d in the `finally` block, and its state was
only ever in memory. Consequence: **no run that has already been paid for can ever be
re-graded**, at any price short of running it again.

**Evidence.**
- `sb/live/runner.py:535-558` — `main()`'s full argparse surface: `--model --driver
  --seed --start --days --limit --corpus --daily-min --daily-max --urgency-horizon
  --reasoning --timeout`. No output flag.
- `sb/live/runner.py:513-516` — the teardown: `mcp.terminate()`, `store.terminate()`,
  `shutil.rmtree(iso_cwd, ...)`. Nothing is read out of the store first.
- `sb/live/store_app.py:22-27` — the entire run state is four module-level in-memory
  objects (`_events`, `_todos`, `_inbox`, `_warnings`) in a uvicorn child process.
  No file, no database, no dump endpoint (`/state` at `:235-237` is a live read only).
- Capture is done by shell redirection outside the harness: `RUNNING.md:78`
  (`NO_COLOR=1 ./run.sh ... > build/run.log 2>&1`), `BENCHMARK_RESULTS.md:153,157`.
  `.gitignore` line `build/` excludes exactly that directory.
- `past/claude-haiku-4-5.md:7` — "**Run log:** build/run_claude-haiku-4-5_v2.log". That
  file does not exist: `find . -name "*.log" -not -path "./.venv/*" -not -path "./.git/*"`
  returns nothing.
- `past/claude-opus-4-8.md` is 0 bytes (`ls -la past/`) — the manual copy step for that
  run was simply never performed, which is the failure mode a manual pipeline has.
- The two `outputs/*.md` headers carry three lines (model / seed-days-emails-start) and
  no `levers` line, because the levers line at `sb/live/runner.py:406-407` and the
  config stamp at `:530-531` postdate them. So the recorded runs carry no lever
  provenance at all, which is why §2.10 of the brief had to *infer* `daily_max=5` from
  the 57-day span.

**Why it matters for benchmark validity.** The register's sequencing argument
(`docs/benchmark-repair.md:38-45`) says doing O first "means every later grader change
can be re-scored offline against runs already paid for". That is true only of runs paid
for *after* O lands. The three-to-four runs already on disk are pretty-printed prose;
the objects the models created, the ids they stamped, and the store state at each day
boundary are gone and cannot be reconstructed. The phase table should read "runs paid
for from here on", and phase 6 (the re-run) is therefore not optional if any phase-2
grader change is to be compared against a real model rather than the oracle.

**Options.**
1. *A `--out DIR` that writes one JSONL/JSON tree per run.* Everything downstream reads
   files instead of scraping stdout. Cost: a schema decision (phase 1), and a second
   format to keep in sync with the human log.
2. *Persist the store instead of the runner.* Add a store endpoint the runner snapshots
   each day, or make the store write-through to disk. Keeps the runner thin, and gives
   the webapp a natural artifact; but it splits the artifact across two processes and
   the store currently has no notion of a "run".
3. *Keep stdout as the only artifact but make it machine-parseable* (structured lines,
   stable field order) and commit the raw log rather than a hand-pasted excerpt.
   Cheapest, no new format, but keeps `sb.analyze`'s regex-scraping coupling (O-6) and
   still cannot carry the store state (O-3).
4. *Do nothing for existing runs; instrument future ones only.* Honest about what is
   recoverable, but leaves phase 1.5 and phase 2 with no real-model baseline until
   phase 6 is paid for.

**Overlaps with:** C (provenance stamping, the broken `build/run_*.log` evidence link,
the empty `past/claude-opus-4-8.md`), O-3, O-5.

**Open questions.**
- Is the artifact meant to be committed to the repo, or to the `backups` branch, or
  pushed to the webapp DB? The 2026-06-09 data loss note in `CLAUDE.md` argues against
  "the only copy lives in `build/`".
- Should a run artifact embed the corpus, or reference it by hash? See O-3.

---

## O-2 The tool trace drops most tool calls, and the loss rate is model-dependent
Status: open
Severity: blocks-measurement
Cost to verify: free-offline

**What's wrong.** `_parse_stream` accumulates assistant messages into a dict keyed on
the message id, so when the CLI emits one `assistant` event per content block — which is
what a parallel tool-call batch looks like on the wire — every block but the last is
overwritten and lost. The loss is not uniform: it depends on how a given model chunks
its tool calls, so the trace systematically penalises models that batch. Every
retrieval statistic in this project, including the brief's §2.7 and §2.8, is computed
from this trace.

**Evidence.**
- `sb/live/runner.py:180-183`:
  ```python
  elif t == "assistant":
      msg = ev.get("message", {})
      resolved_model = msg.get("model") or resolved_model
      seen[msg.get("id", id(ev))] = msg
  ```
  Last write wins per message id; `:187-194` then iterates only the surviving `msg`.
- Mechanism confirmed against the real parser
  (`/private/tmp/.../scratchpad/collapse.py`, five synthetic streams):
  | case | input | `_parse_stream` returns |
  |---|---|---|
  | 1 | five `get_email` blocks, five events, one message id | `['get_email']` |
  | 2 | the same five blocks in **one** event | five `get_email` |
  | 3 | `tool_use` then `text`, same id | `tools=[]`, text survives — the call vanishes |
  | 4 | six assistant messages with **no** `id` | 4 of 6 (see below) |
  | 5 | six assistant messages with `"id": null` | 1 of 6 |
- Case 4 is a second, independent bug: the fallback key is `id(ev)`, the CPython address
  of a dict that is rebound each loop iteration, so the allocator reuses addresses.
  Reproduced three times, deterministically 4/6 on `.venv/bin/python` (3.13.14):
  ```
  .venv/bin/python -c "import json,sys;sys.path.insert(0,'.');from sb.live.runner import _parse_stream;\
  tu={'type':'tool_use','name':'mcp__secretary__get_email'};\
  print(len(_parse_stream('\n'.join(json.dumps({'type':'assistant','message':{'content':[tu]}}) for _ in range(6)))[1]))"
  → 4
  ```
- **Loss proven from the committed logs, without any assumption about CLI framing.**
  The `actual` field only ever renders objects that already matched an answer-key
  keyword (`sb/grader.py:152,168`), so distinct titles seen there are a strict *lower
  bound* on objects created. Compare that to the `create_event`+`create_todo` calls the
  trace recorded:

  | log | create_* calls in trace | distinct titles in `actual` | ≥ calls missing |
  |---|---|---|---|
  | `outputs/opus.md` | 89 | 71 | — (no proof of loss in aggregate) |
  | `outputs/sonnet.md` | 41 | 70 | ≥ 29 |
  | `past/claude-haiku-4-5.md` | 18 | 68 | ≥ 50 |
  | `past/claude-sonnet-4-5.md` | 20 | 65 | ≥ 45 |

  Per day, counting only titles appearing for the first time, the trace logs fewer
  `create_*` calls than objects that demonstrably appeared on **18/57 opus days,
  29/57 sonnet days, 15/16 haiku days, 16/19 past-sonnet days** — so opus's trace is
  lossy too, just less so.
- The smoking gun for the mechanism: `outputs/sonnet.md`'s per-day `get_email`
  histogram is `{1: 57}` — exactly one surviving `get_email` on every one of 57 days,
  with 1 to 5 emails arriving per day. `outputs/opus.md` is `{1:11, 2:13, 3:12, 4:12,
  5:9}`. `list_new_emails` is `{1: 57}` in both, as expected for a genuinely
  once-per-day call. Totals: sonnet 57 `get_email` for 167 emails, opus 166 —
  reproducing §2.7's numbers and explaining them.
- The codex driver does **not** dedupe: `sb/live/runner.py:355-357` appends every
  `item.completed` tool call. So `claude` and `codex` traces have different fidelity by
  construction and their tool counts are not comparable across drivers.
- `sb/analyze.py:51-58` derives its entire `searched` signal from the string
  `search_inbox` appearing on a day's tools line, i.e. from this trace.

**Why it matters for benchmark validity.** The trace is the only record of what the
model *did* as opposed to what survived in the store, and it undercounts by an unknown,
model-dependent amount. Concretely, this undercuts brief §2.8: "`search_inbox` was used
on 1 of 57 days by opus, 1 of 57 by sonnet, 0 of 16 by haiku" is a *lower bound*, not a
measurement — a `search_inbox` issued in the same message as another call is dropped,
and the demonstrated drop rate for `create_*` on those same runs is 41–74%. V's central
claim rests on a number this parser cannot produce. It also means "sonnet only read 57
of 167 emails" must not be reported as a behavioural finding.

**Options.**
1. *Persist the raw CLI stdout per turn and parse it as a separate offline step.* The
   trace stops being a lossy summary computed at runtime; existing runs stay lost.
   Cost: raw stream-json is large and contains full model text.
2. *Fix the accumulator to append per block rather than keyed-overwrite* (and drop the
   `id(ev)` fallback). Small and local, but still leaves the trace as the only record
   and still cannot be applied retroactively.
3. *Move the tool-call record to the server side* — log every MCP/store call with its
   arguments. Driver-independent and immune to CLI framing changes, but only sees the
   benchmark's own tools, not `ToolSearch`/built-ins (see §2.12), and touches the tool
   surface the models are measured through. Cf. O-4.
4. *Both 2 and 3, and cross-check them* — the disagreement between the two is itself
   the measurement of how lossy the client-side parse is.

**Overlaps with:** V (the retrieval-span claim depends on this signal), O-4, O-6, C.

**Open questions.**
- Does the current CLI (2.1.233) actually emit one `assistant` event per content block,
  or does it batch? The log evidence is consistent with per-block emission but does not
  prove it; confirming needs one captured raw stream, obtainable from a bounded smoke.
- Does anything else key on `_parse_stream`'s tool list besides display and
  `sb.analyze`? (`is_error` and `session_id` are load-bearing; tools appear not to be.)

---

## O-3 A final-state dump is not sufficient: the store records no history and no day
Status: open
Severity: blocks-measurement
Cost to verify: free-offline

**What's wrong.** Grading is done incrementally, once per simulated day, against the
node state *as it stood on that day*, so replaying a run offline requires the state at
every day boundary — not the state at the end. The store keeps no creation day, no
mutation history, and physically removes deleted objects, so the day-by-day sequence
cannot be reconstructed from a final snapshot. "Dump the store at the end" therefore
does not make a run re-gradeable.

**Evidence.**
- Grading is per-day and stateful: `sb/live/runner.py:431` captures `before =
  _all_ids(state)`; `:500-501` computes `day_new = _all_ids(state) - before`; `:507-510`
  grades each email against `_node_state(...)` and `_turn_delta(corpus, state, eid_new)`.
- Store objects carry no creation day: `sb/live/store_app.py:32-51` (`EventIn` =
  `email_id, title, start, end, description`; `TodoIn` = `email_id, title, due_date,
  description`) and `:130` / `:160` store exactly `{"id": eid, **e.model_dump()}`.
- Updates overwrite in place with no history: `:143` `_events[eid].update(...)`,
  `:173` `_todos[tid].update(...)`. A title that matched a keyword on day 10 and was
  renamed on day 40 grades differently on day 10 than a final dump would say.
- Deletes are destructive: `:149` `_events.pop(eid, None)`, `:180` `_todos.pop(...)`.
  A `cancel` op grades on absence (`sb/grader.py:155-160`), so the object that was
  cancelled leaves no trace at all.
- What *is* deterministic and therefore need not be captured object-by-object: the
  serve plan. `sb/scheduler.py:15` — "Same (corpus, start_date, seed, levers) =>
  identical plan"; `build_plan` at `:58-156` uses only `random.Random(seed)`. Rendered
  bodies likewise (`sb/live/runner.py:421` renders from `corpus` + `plan.anchors`). So
  `(corpus content or hash, start, seed, days, levers)` + per-day store snapshots is a
  sufficient set; the corpus reference is load-bearing because phase 4 will change it.
- Note for whoever designs the capture: `_node_state`'s fourth parameter `sid_filter` is
  accepted and never referenced in the body (`sb/live/runner.py:137-146`; confirmed by
  AST walk — `'sid_filter' in names` is `False`). The state the grader actually sees is
  the **full cumulative node pool**, and `eid_new` only reaches `TurnDelta`. Diagnosing
  whether that is intended belongs to G/A; it matters here because it fixes what a
  snapshot must contain.
- There is no corpus-hash code anywhere in the repo (`grep -rn "hashlib\|blake2\|digest"
  sb/ scripts/ run.sh` → no matches), yet `past/claude-haiku-4-5.md:4` records
  "sha d737d44e14dc7d20" and `past/claude-sonnet-4-5.md:4` "sha 809d389794dd79a9". Those
  hashes cannot be recomputed with anything in the tree, so they cannot be checked
  against today's `corpus/`.

**Why it matters for benchmark validity.** This is the finding that decides whether
phase 1 actually delivers what the sequencing argument promises. If the capture is a
final-state dump, phase 2 can only re-grade the *last* day correctly and every earlier
day is graded against a future state — silently, with plausible-looking numbers. If the
capture is per-day snapshots (or a day-stamped mutation log), a grader change is exactly
re-scorable. Nothing in the current design makes this impossible; it makes it *absent*,
and the absence is easy to under-fix.

**Options.**
1. *Runner-side per-day snapshot* — the runner already GETs `/state` once per day
   (`:500`); write that response to the artifact along with `before`. Zero store
   changes, but records only what the grader happened to look at, at the moment it
   looked, and misses intra-day ordering and anything deleted mid-day.
2. *Store-side append-only mutation log* — every create/patch/delete recorded with the
   simulated day and a monotonic sequence number. Strictly richer (any day's snapshot is
   derivable, deletions and renames survive, retries become visible), but is a store
   schema change and needs the `/day` value threaded into every write path.
3. *Add `created_on` / `updated_on` to the store objects and dump once at the end.*
   Smallest change, but still loses anything deleted and any pre-rename title.
4. *Capture the corpus itself into the artifact rather than a hash.* Makes a run
   self-contained across phase 4's corpus edits at the cost of ~duplicating 15 node
   files per run.

**Overlaps with:** O-1, A (`email_id` attribution is what a snapshot must preserve),
G (the cumulative-pool question), C (corpus hash / provenance), K (phase 4 corpus edits
invalidate plan-derived replay).

**Open questions.**
- Does re-grading need to replay day by day at all, or is it acceptable to define the
  new grader as a function of the final state plus a mutation log? That is a phase-2
  contract question, not a phase-1 one, but the capture must not foreclose it.
- Should the artifact record the store's `/warnings` list (currently fetched at `:414`
  and `:489` and then discarded at teardown)?

---

## O-4 Retrieval is unobservable: the store logs no reads, so only the lossy trace sees it
Status: open
Severity: blocks-measurement
Cost to verify: free-offline

**What's wrong.** `search_inbox`, `list_new_emails` and `get_email` are read-only store
endpoints that record nothing — no counter, no log line, no entry in `_warnings` except
two coarse heuristics for over-broad searches. The store is the one durable, non-lossy
observer in the system, and it is blind to exactly the behaviour the benchmark exists to
measure. This is the one item in category O that a state dump alone cannot fix: retrieval
leaves no server-side residue, so with the current design a perfectly captured run still
cannot answer "did the model search, and for what".

**Evidence.**
- `sb/live/store_app.py:192-213` `/inbox/search`: the only writes are
  `_warn("broad_search", ...)` when neither `q` nor `sender` is given (`:197-198`) and
  `_warn("large_search_result", ...)` when `limit == 0 or limit > 25` (`:199-200`). A
  normal, targeted search — the behaviour the benchmark wants — is recorded nowhere.
- `sb/live/store_app.py:216-222` `/inbox/new` and `:225-230` `/inbox/{email_id}`: pure
  reads, no `_warn`, no counter.
- By contrast every *write* path is watched: `:128` and `:158` call
  `_watch_attribution` before creating.
- Consequence, measured: across all four committed logs, the only evidence that any
  model ever searched is the tools line — 1/57 days (opus), 1/57 (sonnet), 0/16 (haiku),
  1/19 (past sonnet). Reproduces brief §2.8 exactly, and is subject to O-2's undercount.
- `sb/live/mcp_app.py:112-134` shows the MCP layer is a thin `_call` pass-through with no
  logging either, so instrumenting either layer is a code change to the model-facing
  surface — the thing phase 0 deliberately declined to touch when it refused the
  `mcp` 2.0 port (`docs/benchmark-repair.md:111-114`).

**Why it matters for benchmark validity.** The stated research contribution is retrieval
span (`docs/benchmark-repair-evidence.md:26-27`). The project's headline observation
about that contribution — that retrieval is barely exercised — is derived from a signal
with a proven, unquantified undercount and no independent corroborating source. Until
retrieval is observed somewhere lossless, no claim about search behaviour, in either
direction, is supportable.

**Options.**
1. *Log every store read* (endpoint, query params, result count, simulated day) into a
   structure the runner snapshots. Complete and driver-independent; adds a write path to
   a read endpoint and grows the store's memory footprint over a 167-email run.
2. *Log at the MCP layer instead* (`_call` in `mcp_app.py:32-36`). Catches exactly the
   tools the model was given and keeps the store pure; but the MCP server is a separate
   process whose output currently goes to `DEVNULL`/`PIPE` (`runner.py:242`), so it
   needs its own artifact path.
3. *Count only, don't log arguments.* A per-day counter per tool is enough to fix the
   `searched` boolean and to cross-check O-2's parser, with a much smaller artifact and
   no risk of storing model-authored query text. Loses the ability to ask *what* was
   searched for, which is the more interesting question for V.
4. *Leave the store alone and rely on a fixed client-side parse (O-2 option 2).* No
   change to the measured surface at all, at the cost of remaining hostage to CLI output
   framing.

**Overlaps with:** O-2, V (span / search-rate claims), A (`_warnings` is the existing
precedent for monitor-only instrumentation).

**Open questions.**
- Does adding logging to a read endpoint count as changing the benchmark? Response
  latency changes are the only model-visible effect, and it is small — but the project's
  own rule (`docs/benchmark-repair.md:110-114`) is conservative about the tool surface.
- `ToolSearch` appears in the traces (§2.12) but is a CLI built-in the store never sees.
  Any server-side scheme is blind to built-ins by construction; is that acceptable?

---

## O-5 The log renders only keyword-matched objects, so the dominant failure is unfalsifiable
Status: open
Severity: blocks-measurement
Cost to verify: free-offline

**What's wrong.** `_grade_op` builds its `actual` string from `title_set`, the objects
that already matched the answer key's keywords, and prints `(nothing matching created)`
when that set is empty. So for the failure mode that accounts for ~85% of all failures,
the artifact shows literally nothing about what the model did. Combined with O-1 there
is no surviving record from which a human could tell under-action from a title mismatch,
a kind mismatch, or a wrong-node stamp.

**Evidence.**
- `sb/grader.py:151-152` — `pool = state.events if op.kind == "event" else state.todos`;
  `title_set = [o for o in pool if _title_hit(o, op.match)]`. The pool is filtered by
  kind *and* by keyword before anything is rendered.
- `sb/grader.py:168` — `actual = "; ".join(_fmt_obj(o) for o in title_set) if title_set
  else "(nothing matching created)"`. Non-matching objects are never formatted.
- `sb/grader.py:190` — the no-action branch renders `turn.events + turn.todos`, i.e. only
  objects stamped with *this* email's id (`runner.py:507`), so over-action stamped onto a
  sibling id prints as `(nothing)`.
- `sb/live/runner.py:80-85` `_print_email` prints only `res.details`, which is exactly
  the above. Nothing else about the store reaches the log.
- Measured, over the four committed logs: the distinct object titles ever visible in an
  `actual` field are 71 (opus), 70 (sonnet), 68 (haiku), 65 (past sonnet) — while the
  models demonstrably created more (O-2's table). Everything above that count is
  invisible in every artifact.

**Why it matters for benchmark validity.** Two register items depend on this being
fixable. Brief §4.3 says the split of the 52 "nothing matching created" cases is
*unmeasured*; the stronger and correct statement is that it is **unmeasurable from any
surviving artifact** — the information was never rendered and the state was discarded.
And `docs/benchmark-repair.md:51` schedules phase 1.5, "hand-grade ~30 emails", with cost
`free`. A human cannot hand-grade a model's behaviour they cannot see. Either phase 1.5
is redefined as hand-checking *answer keys against email prose* (genuinely free, and a
different exercise), or it is a paid activity that must run after O lands and after a
fresh instrumented run.

**Options.**
1. *Render the unmatched pool in the artifact* — every object in the node, and every
   object stamped to this email regardless of node/kind, alongside the matched set.
   Makes the four-way split directly countable; makes the human log much noisier, so it
   probably belongs in the machine artifact only.
2. *Emit per-op diagnostic codes* (`kind_mismatch`, `wrong_node_stamp`,
   `nothing_created`) computed by the grader itself. Compact and directly countable, but
   pre-commits to a taxonomy that is a phase-2 decision.
3. *Dump raw store state only, and compute all diagnostics offline.* No grader change at
   all — the grader/corpus separation rule in `CLAUDE.md` stays clean — but every
   consumer re-implements the join.
4. *Leave the grader untouched and accept that §4.3 stays open until phase 6.*

**Overlaps with:** G (what `actual` should mean is a grader-identity question), A
(sibling-stamp invisibility), O-1, O-3.

**Open questions.**
- Does adding fields to `actual` count as "changing the grader" under the rule that
  grader and corpus never move in the same commit? Rendering-only changes cannot move a
  score, and a test asserting score-invariance would prove it.
- How many of the 52 opus cases are recoverable at all? On this analysis: none, from
  existing files.

---

## O-6 Tools and narration are attributed to a day, never to an email, and the narration is truncated
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** One day is one CLI turn, so the runner prints one tools line and one
narration snippet per day and every email that arrived that day inherits both. The
narration — the only per-email statement of what the model believed it did — is cut at
200 characters, which in practice is every day of every run. `sb.analyze` then
reconstructs its dataset by regex-scraping this human log, inheriting the day-level
granularity as a documented compromise.

**Evidence.**
- `sb/live/runner.py:57-66` `_print_day`, whose own docstring says "tools/said belong to
  the whole day's turn, not any one email, so they're reported here". `:65` —
  `snippet = said[:200].replace("\n", " ")`.
- `sb/analyze.py:36-41` — "`searched` therefore means 'the model used `search_inbox` at
  least once during this email's day.'"
- `sb/analyze.py:25-27` — the dataset is rebuilt with three regexes over ANSI-stripped
  text: `_LINE`, `_DAY`, `_ANSI`.
- Measured truncation: narration snippets sitting at the 200-char cap —
  **57/57 days (opus), 50/57 (sonnet), 16/16 (haiku), 18/19 (past sonnet)**. Opus and
  haiku are truncated on literally every day of their runs.
- Measured inflation from day-level attribution: opus's single day-1 `search_inbox` tags
  **5 of 167 emails (3.0%)** as `searched=True`; past-sonnet's single day-1 search tags
  8/176 (4.5%). With haiku's `daily_max=21`, one search on a 21-email day would tag 21
  emails.
- Latent parser fragility: `sb/analyze.py:51` matches with `if "tools" in line`, before
  the PASS/FAIL check, so any line containing the substring resets `day_searched`. Not
  observed — `grep -n "tools"` on all four logs finds it only on genuine tools lines
  (0 other occurrences in each) — but a model narrating the word "tools" would silently
  corrupt the search-rate column.

**Why it matters for benchmark validity.** The search rate is one of the two numbers the
paper's construct-validity claim rests on, and it is computed at the wrong granularity
in a direction that inflates it. The truncated narration is the only place a reader can
check the grader's verdict against the model's own account of the day; at 200 characters
on a 21-email day it cannot serve that purpose.

**Options.**
1. *Record tool calls with their turn/day index and their arguments in the machine
   artifact* and let consumers choose granularity, leaving the human log as is.
   Depends on O-2 or O-4 for a non-lossy source.
2. *Keep full narration in the artifact and truncate only in the terminal.* Trivial, and
   makes narration usable as a cross-check; the artifact then contains the model's full
   prose, which is large and needs a retention decision.
3. *Attribute tool calls to emails by inspecting arguments* — a `get_email(X)` or a
   `create_event(email_id=X)` names its email. Gives genuine per-email attribution
   without changing the day loop, but only for tools that carry an id; `search_inbox`
   carries none, which is the one that matters for V.
4. *Retire `sb.analyze`'s log-scraping and point it at the artifact.* Removes a whole
   class of fragility, at the cost of breaking its ability to read the four historical
   logs — which is the only thing it can currently read.

**Overlaps with:** O-2, O-4, V (search rate, tier reporting), A.

**Open questions.**
- Is per-email tool attribution even well defined under a day loop, given the model may
  interleave? A sequence number plus the day may be the honest answer.
- `sb.analyze` never reads `email.tier` (§2.9); does it get rewritten in phase 1 against
  the new artifact or in phase 5 with V?

---

## O-7 Infrastructure errors and retries are folded into the score with no machine-readable marker
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** A day whose turn fails after seven attempts marks each of its emails
`None`, prints a line the analysis regex does not match, and counts them against the
denominator — so a network failure and a model failure produce the same score and the
same red dot. Separately, the retry loop re-runs the whole day's prompt against a store
that already contains the failed attempt's objects, with no rollback, and the `before`
watermark is taken once outside the loop, so every attempt's objects land in the same
day's delta.

**Evidence.**
- `sb/live/runner.py:431` — `before = _all_ids(...)` is captured **before** the retry
  loop at `:443-486`; there is no rollback anywhere in `sb/live/store_app.py` other than
  the whole-run `/reset` at `:103-112`.
- `sb/live/runner.py:491-496` — on failure, `results[eid] = None` and
  `print(f"[{email_no:>2}] {eid:24s}  ERROR after retries: {detail}")`. That line has no
  `PASS`/`FAIL` token, so `sb/analyze.py:25`'s `_LINE` regex never matches it and the
  email vanishes from the analysis rather than being counted as an error.
- `sb/live/runner.py:518-528` — `errored = len(order) - len(graded)`;
  `pct = passed / len(order)`; the dot bar is
  `f"{GREEN}●{RESET}" if (r and r.passed) else f"{RED}●{RESET}"`, so an errored email is
  rendered identically to a failed one. The `N errored` suffix is printed only when
  non-zero, which is why a reader of a clean log cannot tell the field is even there.
- Retries and usage-limit pauses do print (`:478-479`, `:485`), so they would appear in a
  redirected log — but all four committed logs contain zero `retry` / `session limit` /
  `ERROR after` lines (`grep -nE "retry|session limit|ERROR after"`), so this is latent
  on the current evidence, not observed. `past/claude-haiku-4-5.md:6` records
  "ERROR 0 (clean retry — supersedes the earlier rate-limited 84/176 run)", i.e. the
  project has already had at least one run spoiled this way.
- A duplicate created by a retried attempt fails the uniqueness rule at
  `sb/grader.py:164` (`count_ok = len(title_set) == 1`) with nothing in the artifact to
  attribute it to the retry rather than to the model.

**Why it matters for benchmark validity.** A headline score that silently mixes harness
failures with model failures is not a capability measurement, and there is currently no
field anywhere that separates them. Retry residue is worse: it manufactures exactly the
"found 2 matching, expected exactly 1" failure that §2.3 counts 14 times for opus, and
nothing distinguishes a model that double-booked from a day that ran twice.

**Options.**
1. *Record per-email status as a tri-state* (`pass` / `fail` / `error`) plus a per-day
   attempt count in the artifact, and report the score with and without errored emails.
   Reporting-only; does not fix retry residue.
2. *Snapshot the store before each attempt and roll back before retrying.* Removes the
   residue at the cost of a store rollback endpoint and a decision about what a
   partially-completed turn means.
3. *Refuse to score a run with any errored day* — fail loudly rather than reporting a
   depressed number. Simple and safe; wastes a paid run on one transient blip.
4. *Recompute `before` inside the retry loop* so only the successful attempt's objects
   count. Cheapest, but silently strands the failed attempt's objects in the store where
   they will collide with later days' pools (`sb/grader.py:151`, cumulative).

**Overlaps with:** A (day attribution), G (uniqueness rule), C (a run's status belongs in
the provenance stamp), O-1.

**Open questions.**
- Are the four committed runs genuinely retry-free, or were retry lines dropped when the
  stdout was hand-copied into markdown? Unanswerable — see O-1.
- Should a usage-limit pause (which is not a failure) be distinguished from a transient
  error in the artifact? Currently both only print.

---

## O-8 No cost, timing, token, or version capture — although the CLI hands them over
Status: open
Severity: slows-work
Cost to verify: free-offline

**What's wrong.** The runner reads the CLI's `result` event for two fields and discards
the rest of it, including the cost and duration the CLI reports for the turn. Nothing
anywhere records wall-clock time, token counts, the CLI version, or the `sb` commit, so
no artifact can say what a run cost or how long it took.

**Evidence.**
- `sb/live/runner.py:184-186` — the whole `result` branch:
  ```python
  elif t == "result":
      session_id = session_id or ev.get("session_id")
      is_error = bool(ev.get("is_error"))
  ```
  The event object is otherwise dropped.
- `grep -rn "total_cost_usd\|duration_ms\|num_turns\|usage" sb/` returns exactly two
  hits, both inside the rate-limit regex at `sb/live/runner.py:264-267`. Nothing
  accounts for cost or tokens.
- `sb/live/runner.py:403-408` — the run header prints model, driver, seed, day count,
  email count, start date, levers, corpus path. No timestamp, no CLI version, no repo
  sha. `:530-531`'s config stamp repeats the same set.
- The committed artifacts confirm the gap: run duration and date appear only as
  hand-typed prose (`past/claude-haiku-4-5.md:7`, "completed 2026-07-04"), and
  `outputs/*.md` carry no date at all.

**Why it matters for benchmark validity.** The register's entire sequencing argument is
an economic one — "the only way to evaluate a grader change is to pay for a whole new
run" (`docs/benchmark-repair.md:40-41`) — and the project cannot state what a run costs,
so the argument is unquantified. `CLAUDE.md` warns that "a full run is ~57 day-turns"
and that live runs compete with the interactive session for quota; without recorded
per-turn durations there is no basis for bounding a smoke test or for planning phase 6.
Reproducibility also suffers: M-3 was settled against "CLI 2.1.233", a version no
artifact records.

**Options.**
1. *Keep the whole `result` event per turn in the artifact.* Captures cost, duration,
   turns and usage with no field-by-field decision; ties the artifact to the CLI's
   schema, which changes between versions.
2. *Extract a fixed small set* (`total_cost_usd`, `duration_ms`, `num_turns`, input/output
   tokens) into a stable per-turn record. Stable and small, but silently loses fields the
   CLI adds later, and the codex driver's event shape differs.
3. *Stamp environment provenance once per run* — `claude --version` / `codex --version`,
   `git rev-parse HEAD`, UTC start and end, hostname. Cheap, orthogonal to the per-turn
   question, and directly serves M-5 (whether a run launched from inside an agent session
   is equivalent to one from a plain shell).
4. *Time turns in the runner itself* rather than trusting the CLI's numbers. Driver-
   independent and covers retries and pauses, which the CLI's `duration_ms` will not.

**Overlaps with:** C (provenance stamping is the same artifact field), M-5, O-1.

**Open questions.**
- Does `codex exec --json` expose comparable cost/usage fields? Untestable here —
  `codex` is not installed on this machine (`docs/benchmark-repair.md:167-170`).
- Should a subscription-quota run report cost at all, given it is billed as usage rather
  than per-token? Duration and token counts may be the more meaningful budget.

---

## O-9 Objects with an unparseable date are dropped from grading without a trace
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** `_obj_from` returns `None` when neither `start` nor `due_date` parses
as ISO, and both `_node_state` and `_turn_delta` silently skip those records. The store
accepts any string in those fields, so a model that writes a non-ISO date has its object
excluded from the pool it should have joined — and, on a no-action email, is credited
with "correctly took no action". No warning is emitted and no counter is kept.

**Evidence.**
- `sb/live/runner.py:128-134` — `when = _parse_iso(record.get("start") or
  record.get("due_date") or "")`; `if when is None: return None`.
- `sb/live/runner.py:137-146` and `:149-152` — the walrus filters
  `if (o := keep(r, "event"))` / `if ... and (o := _obj_from(r, "event"))` drop the
  `None` with no branch and no logging.
- `sb/live/store_app.py:32-51` — `start`, `end` and `due_date` are plain `str`, with no
  format validation; `:126-131` and `:156-161` store whatever arrives.
- `sb/grader.py:188-189` — the no-action verdict is `created = turn.events + turn.todos`
  then `passed = not created`, so a dropped object turns an over-action into a pass.
- Nothing in `sb/live/store_app.py`'s `_warn` set (`:86-92`, `:197-200`) covers this;
  `sb/live/runner.py:489-490` prints only the store's warnings.

**Why it matters for benchmark validity.** It is a silent, unbounded error term in both
directions: it can convert a correct action into "nothing matching created" (§2.3's
largest bucket) and an over-action into a no-action pass (§2.6's hole). Its size is
unknown and, per O-1 and O-5, unmeasurable from any existing artifact.

**Options.**
1. *Validate at the store's write path* and reject a non-ISO date with a 4xx the model
   sees. Cleanest signal, but changes the tool surface the models are measured through
   and would change scores.
2. *Warn, don't reject* — emit a monitor-only `_warn("unparseable_date", ...)`, matching
   the existing precedent at `store_app.py:86-92`. Zero score impact, quantifies the
   term, and the warning already reaches the log via `runner.py:489-490`.
3. *Count the drops in the runner* and surface them in the artifact. No store change at
   all, but only visible where an artifact exists (i.e. after O-1).
4. *Do nothing until it is shown to be non-zero*, which requires one instrumented run.

**Overlaps with:** A (no-action grading), G (pool membership), O-5.

**Open questions.**
- Did any of the four recorded runs actually hit this? Unanswerable from the artifacts —
  a dropped object appears in the log exactly as if it were never created.
- Would a unit test on `_obj_from` plus a synthetic store record settle the mechanism
  cheaply enough to skip the live check? (Yes for the mechanism; no for the frequency.)

---

## Notes back to the evidence brief

**§2 items reproduced.** §2.7's tool counts (sonnet 57 `get_email`, opus 166 for 167
emails) reproduce exactly, as does §2.8's search usage (1/57 opus, 1/57 sonnet, 0/16
haiku), §2.6's single `stale_email_id` warning (`outputs/opus.md:67`, the only one in any
log), §2.9's description of the runner's output, and §2.10's oracle-clean corpus.

**§2 items I could not reproduce, or that are stated imprecisely.**

1. **§2.1 says "the three recorded runs" but there are four non-empty logs.** §2.7's
   own evidence — "sonnet's day 1 logs ... for 8 emails" — comes from
   `past/claude-sonnet-4-5.md` (the retired 176-email corpus, 102/176), which is not in
   the §2.1 table. §2.12 does say "four records". The table should list it or the prose
   should stop citing it.
2. **§2.7's quoted trace is missing its first entry.** `past/claude-sonnet-4-5.md`'s
   day-1 tools line reads `ToolSearch, list_new_emails, get_email, search_inbox,
   search_inbox, create_event`, not `list_new_emails, get_email, ...`.
3. **§2.7's conclusion is understated.** It presents the loss as an explanation for a
   sonnet-vs-opus discrepancy; the trace is lossy for **every** model, opus included
   (18/57 opus days log fewer `create_*` calls than objects that provably appeared). See
   O-2.
4. **Several `runner.py` anchors point at the wrong lines**, even though the brief
   (`67b3005`) postdates the last change to that file (`3956826`):
   | brief | what is actually there | correct anchor |
   |---|---|---|
   | §2.6 `runner.py:472-475` "grades a day's objects by splitting on `email_id`" | the retry error/pause code | `runner.py:498-510` |
   | §2.12 `runner.py:308-310` "`--permission-mode bypassPermissions`" | the `build_cmd` signature | `runner.py:320-322` |
  | §2.4 `grader.py:163-165` "`pool` is all objects of that kind" | `matched`/`count_ok` | `pool` is `grader.py:151`; `count_ok` is `:164` |
   | §3 `runner.py:394-477` "day loop, attribution split" | `order = [...]` flattening | day loop `runner.py:412-512` |
   | §3 `runner.py:180` "lossy tool parse" | `elif t == "assistant":` | the overwrite is `runner.py:183` |
   | §2.5 `oracle.py:51` | `for op in email.answer.ops:` | `oracle.py:52` |
   `store_app.py:86-92`, `analyze.py:25`, `grader.py:68-70`, `schema.py:155`,
   `span.py:26-41` and `scale.py:67-96` all check out.

**§4 items I believe are now established.**

- **§4.3 ("the split of the 52 'nothing matching created' cases is unmeasured") should be
  upgraded to *unmeasurable from existing artifacts*.** `sb/grader.py:152,168` renders
  only keyword-matched objects and the store state was discarded (`runner.py:513-516`),
  so no surviving file distinguishes under-action from a title, kind, or node-stamp
  mismatch. Measuring it requires a fresh instrumented run. See O-5.
- **`docs/benchmark-repair.md:51`'s phase 1.5 is not free** for the same reason. A human
  cannot hand-grade model behaviour that no artifact records. Phase 1.5 as scheduled can
  only hand-check answer keys against email prose — a worthwhile but different exercise
  — unless it runs after O and after a paid instrumented run.
- **§2.8's search-usage counts are lower bounds, not measurements**, because they are
  read off the trace that O-2 proves drops 41–74% of `create_*` calls on those same runs.
  The offline span measurement (mean 31.6, max 83) is independent of this and stands; the
  claim that models *do not search* does not.
- **The register's phase-1 payoff sentence is too strong.** "Every later grader change
  can be re-scored offline against runs already paid for"
  (`docs/benchmark-repair.md:40-43`) holds only for runs paid for after O lands. Nothing
  currently on disk is re-gradeable. See O-1.

**Answering the specific question — what makes full offline re-grading impossible rather
than merely absent.** For a *future* run, a sufficient capture is: the corpus (content or
a verifiable hash) plus `start`/`seed`/`days`/levers — from which the plan and every
rendered body are deterministic (`sb/scheduler.py:15`) — plus the store state at every
day boundary with its `email_id` stamps, plus the `before` watermark or an equivalent
day-stamped mutation log. Everything in that list is absent but obtainable. The single
item that is *impossible* under the current design, not merely absent, is **retrieval
behaviour**: reads leave no server-side residue at all (`store_app.py:192-230`), so the
only witness is the provably lossy client-side parse. Recovering it requires
instrumenting a read path — the store or the MCP layer — which is a change to the surface
the models are measured through, not an output flag. That decision cannot be deferred to
"just dump the state".
