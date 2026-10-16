# baseline-sonnet-4-5-r2

The approved rerun of the 1c baseline (`docs/benchmark-repair.md`, roadmap decision 4), run
after the O-7 fix. Same model, seed, levers, corpus and system prompt as
`captures/baseline-sonnet-4-5`. It is the first clean 57-day run and the first same-seed
replicate in this project.

| | |
|---|---|
| requested | `claude-sonnet-4-5` |
| **served** | `claude-sonnet-4-5-20250929`, observed from the CLI's init event every turn |
| score | **123/167 (74%)** as run; every day valid |
| config | seed 42 · start 2026-06-01 · days 60 · daily 1-5 · urgency 7 |
| corpus | `corpus/` sha `03e0d963b9866d8f`, same as 1c |
| harness | Claude Code CLI **2.1.296** (1c: 2.1.235) · 57 turns · 1,423 s |
| validity | 57/57 days `ok`, secretary MCP `connected` on every turn, 1 attempt per day, 0 MCP restarts, 0 errored |
| cost | $7.48 API-equivalent (1c: $7.64) |

## Re-grade it

    python -m sb.regrade captures/baseline-sonnet-4-5-r2

Prints `RESCORE 123/167` and `VALID 123/167 every day valid`, identical to the as-run score.
`sb.regrade` refuses a corpus that changes what the model saw or when it saw it, so a key
repair can be scored here and a prose or plan repair cannot (register C-11);
`--corpus captures/baseline-sonnet-4-5-r2/corpus` grades with the run's own keys.

## Against 1c

Both runs graded by the current grader (`sb.regrade`):

| | 1c | r2 |
|---|---|---|
| all 167 emails | 114 | 123 |
| days 1-50 (148 emails, valid in both) | 110 | 113 |
| days 51-57 (19 emails; 1c had no tools) | 4 | 10 |

On the 148 emails valid in both runs, 141 verdicts agree (95%): 5 went from fail to pass and
2 from pass to fail. The +9 total splits into +6 on the outage tail and +3 on the shared
days. The flips mix run-to-run variance with the CLI version change, so they bound the noise
rather than measure it.

## Against the phase 1d audit

The phase 1d audit labelled 75 checks of the 1c run (`docs/_repair/agent-audit-continuation/RESULTS.md`).
Each label implies what a second run should do. Same grader, check by check:

| 1c audit label | checks | r2 | what the label predicts |
|---|---:|---|---|
| PASS (sampled passes) | 15 | 15 pass | should hold |
| PLAN (delivered too late or impossible) | 10 | 10 fail | cannot be passed |
| GRADER (grader rejects correct work) | 4 | 4 fail | fails until the grader changes |
| KEY (answer key wrong) | 28 | 23 fail, 5 pass | mostly fails |
| MODEL (genuine model error) | 5 | 4 fail, 1 pass | a real weakness repeats |
| AMBIGUOUS, tools down (1c days 51-57) | 8 | 6 pass, 2 fail | unknown; now judgeable |
| AMBIGUOUS, unclear instruction | 5 | 2 pass, 3 fail | a coin flip |

The labels held: 37 of the 42 benchmark-defect checks failed again, every sampled pass held,
and 4 of 5 model errors repeated. The 5 KEY checks that passed (#6, #16, #34, #43, #47) look
like keys that demand more than the email supports, which the model sometimes meets anyway.
The two tools-down checks that still fail with tools (#54 `launch-meeting-conflict-2`, #59
`launch-meeting-recap`) are now unexplained failures. MODEL #13 passed this time while the
same email's other check failed, so the model's behavior on that email varies.

This run fails 51 of 190 checks (1c: 60). By the 1c labels: 37 known benchmark defects,
4 known model errors, 5 ambiguous, and **5 not audited**, all of them 1c passes:

| email | check | r2 |
|---|---|---|
| `Partnership-with-deeptech-companies.caltech-conference-invitation` | create FBS Conference | not created |
| `Innovation-comp.sponsor-call-no-longer-needed` | cancel sponsor call | still on the calendar |
| `shoe-product-launch-delays.design-delays-internal-conflicts` | create Meet with Melissa 1 | not created |
| `World_Cup_Cleat_Launch.reveal-event-date-and-venue-2` | create Reveal Rehearsal | wrong day |
| `Sponsoring-Marathon.launch-meeting-conflict` (day 51) | no action | over-acted |

The last one is one of 1c's four outage-day no-action passes: with tools, the model acted on
it, so that 1c pass was earned only because the tools were down.

## What this run shows about O-7

The MCP server's stderr went to `mcp_server.log` (81,944 bytes by the end). Leaving out the
runner's own health checks (102 `tools/list` requests by then, 7,140 bytes), the model's and
CLI's traffic had written 65,307 bytes when day 51's turn opened: 229 bytes short of the
64 KiB at which the pre-O-7 server blocked, and 1c lost its tools during day 51. Each Claude Code turn opens a session and
lists tools, prompts and resources before its first tool call (57 of each `prompts/list` and
`resources/list` over 57 turns; 344 tool calls in all). The run then continued past that
point with no hang.

## Cost field

CLI 2.1.296 reports `total_cost_usd` cumulatively over the resumed session (it rises every
turn: $0.18 on day 1, $7.48 on day 57). 2.1.235 reported it per turn. The run's cost is
therefore the day 57 value, not the sum over `raw/` (which gives $183 here).

## Files

As in `captures/baseline-sonnet-4-5`, plus:

| path | what |
|---|---|
| `mcp_server.log` | the secretary MCP server's stderr (O-7) |
| `days/NNN.json` | also `mcp_status` and `attempts` per day |
| `manifest.json` | also `errored_days`, `retried_days`, `mcp_restarts` |
| `corpus/nodes/*.json` | frozen copy of the corpus the run was served (sha `03e0d963b9866d8f`); added 2026-10-16 from the unchanged `corpus/` (C-11) |

No `raw/day_NNN.attempt_K.jsonl` files: no attempt failed.

## Redaction

Before committing, the repository path was replaced with `<repo>` and the home-directory
path with `/Users/REDACTED`: once per raw stream (the CLI's `memory_paths.auto` field), once
in `run.log` (the capture path) and once in `mcp_server.log` (a dependency warning). Nothing
else was altered. `sb.regrade` gives 123/167 before and after.
