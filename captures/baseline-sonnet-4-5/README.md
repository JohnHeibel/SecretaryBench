# baseline-sonnet-4-5

Phase 1c of `docs/benchmark-repair.md`. The first run in this project's history whose
served model is **certified rather than asserted**, and the first that can be re-graded
offline without paying for a new run.

| | |
|---|---|
| requested | `claude-sonnet-4-5` |
| **served** | `claude-sonnet-4-5-20250929` — observed from the CLI's init event, every turn |
| score | **97/167 (58%)** as run; re-graded **114/167** all days, **110/148** valid days (see below) |
| config | seed 42 · start 2026-06-01 · days 60 · daily 1-5 · urgency 7 |
| corpus | `corpus/` sha `03e0d963b9866d8f` (`sb.live.runner._corpus_hash`) |
| harness | Claude Code CLI 2.1.235 · 57 turns · 0 errored as recorded, but days 51-57 ran without the secretary tools (below) · no model drift |

## Re-grade it

    python -m sb.regrade captures/baseline-sonnet-4-5

Prints `RESCORE 114/167` (all days) and `VALID 110/148` (days 51-57 excluded) against
`as-run 97/167`; the 17-point gap to the as-run score is the phase 2 grader contract. Any
change to `sb/grader.py` can be scored against this run for free, as many times as you like.
That is the entire point: before this capture existed, evaluating a grader change meant
paying for a whole new run. `sb.regrade` refuses a corpus that changes what the model saw or
when it saw it, so a key repair can be scored here and a prose or plan repair cannot
(register C-11); `--corpus captures/baseline-sonnet-4-5/corpus` grades with the run's own keys.

## Tool outage, days 51-57 (register O-7)

The secretary MCP server stopped answering during day 51 and never came back. Every turn
from day 51 to day 57 started with `mcp_servers: [{name: secretary, status: failed}]` in its
init event (`raw/day_051.jsonl`-`day_057.jsonl`, line 1) and made no tool calls; the model
said "the MCP server has disconnected". The runner of the time did not read that status, so
those 7 days (19 emails, 20 ops) are recorded `ok: true` and graded as if the model had
chosen to do nothing: 16 ops fail and 4 no-action emails pass by default. The "0 errored"
above is what the runner recorded, not what happened.

The cause (an unread stderr pipe that filled and blocked the server) and the runner fix are
in register O-7. This capture is left as it was run: `sb.regrade` finds the outage days from
the raw init events, without editing `days/*.json`, and reports both scores. Neither is "the"
score. The excluded tail is not a random sample: moves and cancels are 6 of its 20 ops
against 26 of 190 overall, so 110/148 is measured on an easier mix.

## Files

| path | what |
|---|---|
| `manifest.json` | config, certified served model, corpus hash, serve plan, score |
| `verdicts.json` | per-email grade and details |
| `days/NNN.json` | per day: full store state, new object ids, attribution map, tools |
| `state_final.json` | every object at end of run |
| `warnings.json` | attribution warnings |
| `raw/day_NNN.jsonl` | the CLI's raw stream, so a later parser fix (O-2) applies here |
| `run.log` | human-readable stdout, as printed |
| `corpus/nodes/*.json` | frozen copy of the corpus the run was served (sha `03e0d963b9866d8f`); added 2026-10-16 from the unchanged `corpus/` (C-11) |

`days/NNN.json` holds the **complete** store state, not only the objects the grader
matched. In this run the model created **81** objects and the printed log renders
**65**: 16 objects (20%) exist nowhere in the human record. That gap is register
finding O-5, and closing it is why this capture exists.

## Redaction

The raw streams and `run.log` were edited in exactly one respect before committing:
the absolute home-directory path was replaced with `/Users/REDACTED`, and
the repository path with `<repo>`. In the raw streams this string
occurred once per file, only inside the CLI's own `memory_paths.auto` field, which
records where Claude Code keeps its session memory and has no bearing on the benchmark.
Total change: 57 bytes across 57 files. Verified by re-running `sb.regrade` after the
edit and confirming 97/167 unchanged. Nothing else was altered.
