# Evidence brief — SecretaryBench repair

Input for the category fan-out (phase A of `docs/benchmark-repair.md`). Every agent
writing a register section receives this file.

**Read this as a starting point to attack, not a conclusion to elaborate.** It is one
investigator's findings from one session. If a claim here is wrong, overstated, or
unfalsifiable, saying so is the most valuable thing you can return. A fan-out that only
confirms this brief has verified nothing; it has laundered one person's assumptions into
a document that looks independently sourced.

Every claim below carries a source. Hold your own findings to the same bar: `file:line`
for code, a quoted line for logs, a command for anything measured.

---

## 1. What the benchmark is

An academic temporal-reasoning benchmark for LLM "secretary" agents. A model plays an
executive assistant over ~2 simulated months, gets a blind prompt each day ("Today is
Jun 4, new mail arrived"), pulls its own inbox through MCP tools, and acts: create event,
create to-do, reschedule, cancel, or correctly do nothing. The grader compares the
resulting calendar state to a hand-authored answer key.

The stated research contribution is **retrieval span**: the fact you need was set weeks
ago and has scrolled out of context, so you must `search_inbox` to recover it.

```
corpus/nodes/*.json  →  sb/scheduler.py  →  sb/live/runner.py  →  sb/grader.py
(authored storylines,   (seeded DAG →       (one CLI subprocess   (state-based,
 emails + date tokens    day-by-day serve     per simulated day)    per obligation)
 + answer-key ops)       plan, deterministic)
```

---

## 2. Established facts

### 2.1 The three recorded runs

| artifact | model | score | days | levers |
|---|---|---|---|---|
| `outputs/opus.md` | claude-opus-4-8 | 90/167 (54%) | 57 | default, `daily_max=5` |
| `outputs/sonnet.md` | claude-sonnet-4-5 | 91/167 (54%) | 57 | default, `daily_max=5` |
| `past/claude-haiku-4-5.md` | claude-haiku-4-5 | 98/167 (59%) | 16 | `daily_max=21` |

Haiku scoring above Opus is the central anomaly. Note Haiku ran a different
configuration, so it is not directly comparable; Opus and Sonnet are.

`past/claude-opus-4-8.md` is **0 bytes** — the Opus run was never written into it.
`BENCHMARK_RESULTS.md` §5 still lists the retired 176-email corpus results (haiku 84/176,
sonnet 102/176) and its two evidence links point at files that do not exist.

### 2.2 These were genuinely three different models

Established forensically, because if all three logs were one model the whole failure
analysis would be analysis of one model. Comparing the titles each run actually produced:

| | em-dash in title | mean title length | byte-identical to opus |
|---|---|---|---|
| opus | **64%** | 46 chars | — |
| sonnet | **0%** | 36 chars | 10% (6/63) |
| haiku | **0%** | 40 chars | 5% |

A single model does not write em-dashes into 64% of its titles in one run and 0% in
another. Opus-vs-sonnet verdict agreement is 148/167 (88.6%), which is high but not
discriminating — see §4.1.

### 2.3 Failure mode tally

Every `why` line across the three logs:

| reason | opus | sonnet | haiku |
|---|---|---|---|
| no object titled like "X" was created | 52 | 47 | 52 |
| found N matching, expected exactly 1 | 14 | 12 | 17 |
| on the wrong day | 18 | 18 | 7 |
| over-acted on a no-action email | 5 | 5 | 0 |
| cancel left something behind | 3 | 4 | 4 |
| *(passing)* correctly took no action | 51 | 51 | 56 |
| *(passing)* matched | 41 | 48 | 49 |
| *(passing)* cancelled | 6 | 5 | 5 |

Roughly 85% of failures are the grader failing to **find** the object. "On the wrong
day" — the only category that unambiguously measures temporal reasoning — is about 10%.

In all 52 opus cases of "no object titled like X", the `actual` field reads
`(nothing matching created)`, meaning nothing in that node of that kind matched the
keyword at all. That is consistent with four different causes and the log cannot
distinguish them: genuine under-action, a title lacking the keyword, a wrong-node
`email_id` stamp, or a kind mismatch (event created where a to-do was keyed).

### 2.4 Grader mechanics

- `grader.py:68-70` `_title_hit` builds the haystack as `f"{obj.title} {obj.description}"`,
  lowercased, and requires **every** keyword to appear as a substring. Descriptions count.
- `grader.py:163-165` `pool` is all objects of that kind in the node, **cumulative over
  the whole run**; `count_ok = len(title_set) == 1`. More emails in a node means more
  chances for a keyword to collide.
- `pool` is filtered by `op.kind`, so an event created where a to-do was keyed produces
  "nothing matching created", indistinguishable from doing nothing.
- `match` defaults to `[name]` (`schema.py:155`).
- `schema.py:551-572` lint check #5 catches author-vs-author keyword containment among
  same-kind `create` ops. It cannot see the titles a model will actually invent.

Observed `match` keywords in the corpus include `["pitch"]`, `["board"]`, `["list"]`,
`["start"]`, `["end"]`, `["ai"]`, `["Team_pizza_party"]`, `["sponsorshippitch"]`.
Note `"ai"` is a substring of "email"; `"end"` of "attend", "recommend", "vendor".

### 2.5 The oracle cannot detect this class of bug

`oracle.py:51`: `title = " ".join(op.match) if op.match else op.name`. The reference model
titles every object using the answer key's own keywords, so it scores 100% on a corpus
whose keywords no real model could match. The oracle validates that an answer is
**satisfiable**, never that it is **gradeable**.

### 2.6 Attribution and no-action

`runner.py:472-475` grades a day's objects by splitting on the `email_id` the model
stamped. For a no-action email the check only inspects objects stamped with *that*
email's id, so over-action stamped onto a sibling id escapes grading entirely.
`store_app.py:86-92` logs `invalid_email_id` / `stale_email_id` as monitor-only warnings.
Across all three runs those warnings fired **once** (opus), so this is a theoretical hole,
not an observed exploit. `docs/POTENTIAL_GAMING.md` anticipates it.

### 2.7 The tool trace is lossy, not merely shallow

`runner.py:180` stores assistant messages into a dict keyed on message id, so multiple
assistant events sharing an id overwrite one another. Evidence: sonnet's day 1 logs
`list_new_emails, get_email, search_inbox, search_inbox, create_event` for 8 emails while
its own narration on the same line says it processed all 8 and created at least 3 events.
Totals across the full runs: sonnet 57 `get_email` calls for 167 emails, opus 166. Same
harness, so those counts are not comparable.

This matters beyond display: `analyze.py`'s entire `searched` signal is whether the
string `search_inbox` appears on that day's tools line.

### 2.8 Retrieval span is barely exercised

`search_inbox` was used on 1 of 57 days by opus, 1 of 57 by sonnet, 0 of 16 by haiku.
The offline span measurement on the authored corpus is mean 31.6, max 83, n=24 needles.
`sb/scale.py` exists to bury needles in filler and force retrieval distance; none of the
recorded runs used it.

### 2.9 No re-gradeable artifacts

The runner only prints; `main()` has no `--out`/`--json` flag. `analyze.py:25` re-parses
the ANSI-stripped human log with a regex. The store is terminated in the `finally` block
and its state is never persisted. **Consequence: every grader change requires a fresh
paid run to evaluate.** `analyze.py` also never reads `email.tier`, so the by-tier report
`TIER_LIST.md` asks for does not exist in code.

### 2.10 Corpus state (measured 2026-08-17)

15 nodes, 167 emails, 134 graded ops. Tiers: 50 T1 / 67 T2 / 50 T3, close to the 30/40/30
target. But `Innovation-comp` alone is 48 emails (29% of the corpus) with only 17 ops and
**zero T3**.

```
.venv/bin/python -m sb.scale --filler 0 --seed 42 --days 200 --dst build/scaled0
→ 167 emails, 167 served over 57 days, oracle: 167/167 = 100%
```

The corpus lints clean and every answer key is satisfiable. The 57-day span matches the
opus and sonnet logs exactly, confirming those ran at default levers, not the
`daily_max=21` pinned in `BENCHMARK_RESULTS.md` §1.

Authoring drift is visible in `corpus/nodes/pizza-party.json` alone: free-typed dates in
bodies ("booked on 6/8", "move the pizza party to June 9") that do not trace to a token
while the answer uses `nth:2,TUE,0m`; a stray `{serve }` rendering a date mid-sentence;
`by tomorrow{!ordering_date = this:FRI}` rendering as "tomorrowFriday, June 12th"; a body
saying "by Friday" with an answer of `by: next:MON`; three emitted anchors nothing
references; a *party* keyed as `kind: todo`.

### 2.11 Environment (all fixed in phase 0, see the register)

`run.sh:26` dropped every CLI argument on the `live` branch. `requirements.txt` pinned
`mcp>=1.0.0` unbounded and mcp 2.0 moved `FastMCP`. macOS `python3` is 3.9.6 and the
harness needs 3.10+. The runner never reported the model the CLI actually served.

### 2.12 Tool surface the model under test sees

`runner.py:308-310` launches the CLI with `--permission-mode bypassPermissions` and no
tool allowlist (`--tools ""` was removed because it zeroed the MCP tools —
`BENCHMARK_RESULTS.md` §0). The model therefore has all built-in CLI tools alongside the
benchmark's MCP tools; `ToolSearch` appears in every run log. The isolated temp cwd blocks
`CLAUDE.md` ingestion but does not block `Read`/`Bash` from reaching
`/Users/jamesoc/dev/SecretaryBench/corpus/nodes/*.json`, which is the answer key. No log
shows any model doing this (no `Bash`, `Read`, `Grep`, `Glob` in any of the four records).

---

## 3. Code anchor map

| what | where |
|---|---|
| object matching, exactly-one rule | `sb/grader.py:68`, `:149-182` |
| no-action grading | `sb/grader.py:186-195` |
| match keyword default, lint #5 | `sb/schema.py:155`, `:551-572` |
| oracle titling | `sb/oracle.py:51` |
| day loop, attribution split | `sb/live/runner.py:394-477` |
| lossy tool parse | `sb/live/runner.py:180` |
| CLI invocation, tool surface | `sb/live/runner.py:296-316` |
| attribution warnings | `sb/live/store_app.py:86-92` |
| log re-parsing, tier unread | `sb/analyze.py:25`, `:74-122` |
| span measurement | `sb/span.py:26-41` |
| filler generation | `sb/scale.py:67-96` |

---

## 4. What is NOT established

Be careful with these. They are the places where a confident-sounding finding is most
likely to be wrong.

**4.1 The two problems mask each other.** The 88.6% opus/sonnet verdict agreement cannot
discriminate "same model" from "different models converging on a broken grader", because
roughly 100 of the 167 outcomes are determined by the grader rather than the model
(51 no-action passes both get right, ~50 keyword failures both get wrong). Do not cite
verdict agreement as evidence for either hypothesis.

**4.2 The true score is unknown.** Nobody has hand-graded a sample. Every claim of the
form "the model was actually right" rests on reading `actual` fields in the log, which
only show objects that already matched a keyword. Objects that matched nothing are
invisible. The register's phase 1.5 exists to fix this.

**4.3 The split of the 52 "nothing matching created" cases is unmeasured.** Under-action
vs title mismatch vs kind mismatch vs wrong-node stamp. Anyone claiming a proportion
needs to say how they measured it.

**4.4 The ~51% run has no artifact.** The reported symptom was "about 51%"; no committed
log shows 51%. At least one run exists whose output was never saved.

**4.5 Phase 0 fixed the model-resolution bug but did not fix the convergence.** Three
genuinely different models still landed within five points of each other with Haiku above
Opus. That is the grader, and it is untouched.

---

## 5. Instructions for fan-out agents

You are writing **one section** of `docs/benchmark-repair.md`. You are read-only on code:
diagnose, do not implement.

**Write your section to its own file** at `docs/_repair/<ID>.md` (create the directory).
Do not return the section as prose — return only a 3-line summary: how many findings, the
highest-severity one, and anything in this brief you believe is wrong.

Fill this template per finding, identically:

```
## <ID>-<n> <short title>
Status: open
Severity: blocks-measurement | distorts-measurement | slows-work | cosmetic
Cost to verify: free-offline | needs-one-live-run | needs-full-roster

What's wrong (3 sentences max)
Evidence (file:line and/or quoted log lines — every claim sourced)
Why it matters for benchmark validity
Options (2-4, with tradeoffs, NO recommendation)
Overlaps with: <other IDs>
Open questions
```

Two rules that matter more than the template:

1. **Diagnose, do not design.** Produce options with tradeoffs. A finished design here
   pre-commits us to a decision we have not thought about, which is the opposite of going
   deep on one problem at a time.
2. **Every claim carries a source.** An unsourced assertion gets cut by the verifier.

Also report, explicitly: anything in §2 you could not reproduce, and anything in §4 you
believe you *have* now established (with the measurement).

### Categories

| ID | scope |
|---|---|
| G | grader identity: object matching, kind filter, description in haystack, cumulative-pool collisions, lint #5's blind spot, the oracle's inability to detect ungradeable keys |
| A | attribution and no-action: `email_id` routing, the sibling-stamp hole, monitor-only warnings |
| O | run artifacts: no machine-readable output, no store dump, lossy tool trace, per-day-not-per-email attribution, truncated narration, no cost/timing/version |
| C | config and provenance: unstamped levers and corpus hash, stale `BENCHMARK_RESULTS.md`, broken evidence links, empty `past/claude-opus-4-8.md` |
| K | corpus authoring: free-typed dates, malformed tokens, prose/answer mismatch, unused anchors, kind choices, `Innovation-comp` dominance |
| V | construct validity: retrieval span near zero, `sb.scale` unused, tier data unread, binary per-email metric |
| M/E | mostly closed in phase 0 — only M-5 (inherited environment) and the doc/setup trail remain |

Then one **synthesizer** agent merges `docs/_repair/*.md` into the register (dedupe,
resolve contradictions, assign permanent IDs, build the overlap map, draft the dashboard)
and one **verifier** agent red-teams the merged result against the code, hunting for
claims that are wrong, unfalsifiable, or overstated.
