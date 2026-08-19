# baseline-sonnet-4-5

Phase 1c of `docs/benchmark-repair.md`. The first run in this project's history whose
served model is **certified rather than asserted**, and the first that can be re-graded
offline without paying for a new run.

| | |
|---|---|
| requested | `claude-sonnet-4-5` |
| **served** | `claude-sonnet-4-5-20250929` — observed from the CLI's init event, every turn |
| score | **97/167 (58%)** |
| config | seed 42 · start 2026-06-01 · days 60 · daily 1-5 · urgency 7 |
| corpus | `corpus/` sha `03e0d963b9866d8f` (`sb.live.runner._corpus_hash`) |
| harness | Claude Code CLI 2.1.235 · 57 turns · 0 errored · no model drift |

## Re-grade it

    python -m sb.regrade captures/baseline-sonnet-4-5

Prints `RESCORE 97/167` against `as-run 97/167`. Any change to `sb/grader.py` can be
scored against this run for free, as many times as you like. That is the entire point:
before this capture existed, evaluating a grader change meant paying for a whole new run.

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

`days/NNN.json` holds the **complete** store state, not only the objects the grader
matched. In this run the model created **81** objects and the printed log renders
**65**: 16 objects (20%) exist nowhere in the human record. That gap is register
finding O-5, and closing it is why this capture exists.

## Redaction

The raw streams and `run.log` were edited in exactly one respect before committing:
the absolute path `/Users/jamesoc` was replaced with `/Users/REDACTED`, and
`/Users/jamesoc/dev/SecretaryBench` with `<repo>`. In the raw streams this string
occurred once per file, only inside the CLI's own `memory_paths.auto` field, which
records where Claude Code keeps its session memory and has no bearing on the benchmark.
Total change: 57 bytes across 57 files. Verified by re-running `sb.regrade` after the
edit and confirming 97/167 unchanged. Nothing else was altered.
