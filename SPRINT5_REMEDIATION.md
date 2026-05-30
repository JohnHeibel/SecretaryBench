# Sprint 5 Remediation Plan

**Branch:** `sprint5-remediation`
**Status:** plan only, no code changed yet
**Audience:** an ultracode session (multi-agent workflow) or a human implementing the fixes
**Source of findings:** a 46-agent adversarial audit of `main` plus first-hand reads of every core file. Every claim below was verified against the source, not just the docs.

> **Scope today: this only drives `claude -p`.** The automated runner supports exactly one harness, Claude Code in print mode (`claude -p`). Both runner implementations build a `claude -p` subprocess (`model_runner.py:345-364`, `harness.py:133-152`), and the only other registered adapter, `CodexAdapter` (`harness.py:182-200`), is a stub that raises `NotImplementedError`. So "drop any harness" is currently a clean interface (`HarnessAdapter`) with exactly one working implementation behind it. Note the two axes are different: swapping the *model* already works (`--model`, and any OpenRouter model via the Anthropic-compatible endpoint), but swapping the *harness* does not. Also note: the MCP server is reachable by any MCP client (see `MCP.md`), but the *scored* 100-day simulation in `engine.py` only knows how to launch and parse `claude -p`. **Section 7 is the recipe for making it truly harness-agnostic.**

---

## Execution kickoff (fresh session: start here)

You are a fresh ultracode session with none of the prior conversation. Do this in order.

**Step 1.** Read Section 1 (the architecture decision) and Section 5 (orchestration).

**Step 2. The four decisions are pre-defaulted below.** These are sensible defaults chosen with the repo owner. If you want to change any, call `AskUserQuestion` now, before writing code. Otherwise proceed with them as written.

- **D1, Architecture: Option A, a `harness/` package.** One shared core, one live path. Create a `harness/` package with `base.py` (the single source of truth: MCP config, hidden-tools set, the full system prompt, `build_user_message`, `build_prompt`, `bootstrap_calendar`, and `parse_stream_output` with token/tool/compaction logging) plus `cli_base.py` / `claude_p.py` (Claude-specific launch and flags). `ClaudeCodeAdapter` becomes a thin consumer of the core. Delete `model_runner.py` as a separate runner, or reduce it to a wrapper that calls the core, so there is exactly one behavior. Harvest structure from the `architecture-update-with-agents-sdk` branch but do not merge it wholesale (it predates the session-continuity, logging, and tool-trimming work on `main`).
- **D2, FIX-5 free-text grading: "ungraded marker" plus fix the splitter.** Genuinely free-text criteria are excluded from `max_score` and reported in a separate `ungraded` list (no silent auto-pass-and-count). AND fix `grader._parse_sub_criteria`, because the real data (Appendix C) shows the splitter currently fragments real prefixed criteria into auto-passing junk: 13 phantom `^` entries, plus `or CC-{...}` and `Update: TC-{...}` fragments whose prefix is no longer at position 0. Those are not free-text, they are real `CC`/`TC` checks the splitter broke.
- **D3, FIX-12 scenario_id: detect-and-log at the MCP server.** Keep "the model must pass `scenario_id`" (it is part of the test), but in the MCP write tools, when a write 404s on a non-existent `scenario_id`, log a distinct line (expected vs seen) and count it, so a garbled id becomes a visible diagnostic instead of a silent 0. Do not auto-correct it.
- **D4, OpenRouter: leave it exactly as is.** Wired, off by default, unverified. Do not build it out, do not remove it (it is intended Sprint 5 scope). Section 7's seam work makes finishing it cheap later.

**Step 3.** Start Phase 0 (Section 5): build the `harness/` package, move the shared logic into it, point the adapter at it. This is sequential and blocks everything else.

**Step 4. Use the `Workflow` tool for fan-out** (ultracode is on). Phase 1's three tracks are file-disjoint and parallelizable; `engine.py` is the one shared file, so route all its edits through a single agent. See Section 5.

### New files you will create (none exist in the repo yet)

- The `harness/` package (`base.py`, `cli_base.py`, `claude_p.py`; `codex.py` optional), per D1.
- `tests/fixtures/sample_stream_json.jsonl`: a recorded `claude -p` stream with a duplicated `message.id`, for the FIX-3 dedup test. Construct it by hand if you have no recording.
- `tests/fixtures/stress_chain.xlsx` (or a generator `tests/generate_stress_chain.py`): a 50+ email single-scenario chain so compaction can actually fire (FIX-10).

### Baseline test reality (read before trusting "tests green")

Current `main` baseline: **130 pass, 6 fail, 1 error.** All seven are pre-existing and unrelated to these fixes:

- **4x `tests/api/test_emails.py::test_delete_email_*`**: stale. They test `DELETE /emails`, a feature deliberately removed (the API returns 405; `README.md:12` says so). Delete or rewrite them. Do NOT implement `DELETE /emails`.
- **2x `tests/test_e2e.py::test_perfect_stub_scores_max` / `test_bad_stub_scores_only_no_action`**: need a live server. They drive `run_simulation` with a stub model (no `claude`), but `pipeline.register_scenario` POSTs to `localhost:8000`, so they pass only with `uvicorn` running.
- **1x `tests/test_pipeline.py::test` (error)**: benign. A helper `def test(name)` at `tests/test_pipeline.py:48` gets collected by pytest as a test ("fixture 'name' not found"). Rename it (e.g. `_case`).

### How to validate your work

- **Offline (no server, no claude):** `python -m pytest tests/ -q -k "not test_perfect_stub and not test_bad_stub"`. Exercises grader, flow controller, parsing, models.
- **Engine integrity (server up, no claude):** start `uvicorn app.main:app`, then `python -m pytest tests/test_e2e.py -q`. The perfect-stub test scoring max is your "engine + pipeline + grader still work end to end" smoke test after the refactor (it uses a stub `model_fn`, so it does not touch the adapter). This is the concrete Phase 0 pass gate.
- **Adapter fixes (server up + `claude` authenticated):** a real run, `python engine.py Emails.xlsx`, is the only way to validate FIX-1 and FIX-3 (calendar events actually created, `token_usage.jsonl` non-empty). There is no offline substitute for the adapter path.

### Adversarial gate (per P0 fix, before calling it done)

After each P0 fix lands, spawn a separate review agent (via `Workflow` or `Agent`) told to *prove the fix is still broken*: for FIX-1, run the bench and confirm an event is actually created on the adapter path; for FIX-2, grade a wrong-date event and confirm it now fails; for FIX-3, confirm `token_usage.jsonl` has rows with cache splits. Only close the fix when the skeptic cannot break it.

### Realistic test Definition of Done

After cleaning the 4 stale delete tests, renaming the `test_pipeline.py:48` helper, and adding the new tests each fix names: **0 failed, 0 errors with `uvicorn` running.** Note that "green" is not achievable on raw `main`; it is achievable only after this cleanup.

---

## 0. How to use this document

This is a work order, not a discussion. Each fix is a self-contained unit with: the problem, the evidence (`file:line` as of the current `main`), the root cause, the concrete change, the files it touches, acceptance criteria, the tests to add or run, and its dependencies. Line numbers are accurate as of commit `89a0689` and may drift one or two lines after edits, so match on surrounding code, not the raw number.

If you are ultracode, read Section 1 (the one decision that shapes everything), then Section 5 (the orchestration plan). The orchestration plan tells you what to fan out, what must be sequential, and where the file conflicts are so parallel agents do not stomp each other.

**Golden rule for this codebase:** the benchmark's job is to measure whether a harness plus a model can run a secretary workload across 100 simulated days. A "fix" is correct only if it makes the *score mean something*. Do not optimize away the token cost of long chains (that cost is the benchmark). Do remove noise, broken paths, and grading that passes regardless of behavior.

---

## 1. The one decision that shapes everything: collapse the two runners

There are two implementations of "drive `claude` for one email":

- `model_runner.py` (`run_model_turn`): complete and correct. Bootstraps a calendar, injects `calendar_id`, full system prompt, full token/tool/compaction logging, session continuity, stream-json dedup.
- `harness.py` (`ClaudeCodeAdapter.run_turn`): the abstraction layer that the engine **prefers by default**. It is a thinner port that dropped the calendar bootstrap, dropped all logging, and ships a system prompt roughly half the size.

The engine's dispatch (`engine.py:362-381`) is `model_fn` then `adapter` then `model_runner` then mock. Because `harness.py` always imports cleanly, **the adapter always wins and `model_runner.py` is dead code on every normal run.** Almost every P0 and P1 bug below is a direct consequence of the adapter being an incomplete copy of the runner.

**Decision: Option A, using a `harness/` package** (pre-defaulted as D1 in the Execution Kickoff; override before Phase 0 only if you disagree). The two options considered were:

- **Option A (recommended): one shared core, one live path.** Create `runner_core.py` (or a `harness/` package) holding the single source of truth for the system prompt, MCP config, hidden-tools list, `build_user_message`, `bootstrap_calendar`, and the stream-json parser (session id + token/tool/compaction logging). `ClaudeCodeAdapter` and any future adapter call into it. Delete `model_runner.py` as a separate runner (or reduce it to a thin wrapper that calls the same core) so there is exactly one behavior. Miguel's unmerged `architecture-update-with-agents-sdk` branch is a prototype of this shape (it adds a `harness/` package and a shared `bootstrap_calendar`), but it predates the session-continuity, logging, and tool-trimming work on `main`, so harvest its structure, do not merge it wholesale.
- **Option B (faster, worse): patch the adapter in place.** Copy the missing pieces (calendar bootstrap, full prompt, logging) from `model_runner.py` into `harness.py`. This fixes the runtime bugs but leaves two diverging copies and the duplication findings (FIX-7) stand. Only choose this if you are time-boxed.

Everything in P0/P1 assumes **Option A**. If you choose B, the same fixes apply, you just apply them inside `harness.py` instead of a shared core.

---

## 2. Priority 0: the default run produces a meaningless score

These three break the benchmark signal on `python engine.py Emails.xlsx` (the default invocation). Nothing else matters until these are fixed.

### FIX-1: Restore calendar bootstrap and `calendar_id` on the live path  `[CRITICAL]`

**Problem.** On the adapter path the model can never create a calendar event, so every `CC` (calendar-created) and `RS` (reschedule) scenario scores 0.

**Evidence.**
- `harness.py:203-216` `_build_user_message(email, sim_date, scenario_id)` has no `calendar_id` and never emits a `calendar_id =` line.
- `harness.py` has zero calendar code. `create_calendar` is a disallowed tool (`harness.py:69-74`) and there is no `list_calendars` tool in `mcp_server/server.py`, so the model cannot create or discover a calendar at runtime.
- `app/routers/calendar.py:88-92` returns 404 on `create_event` if the calendar does not exist.
- The working reference: `model_runner.py:125-151` (`_bootstrap_calendar`), called at `:338`, with `calendar_id` injected into the prompt at `:229`.

**Root cause.** Calendar lifecycle lived only in `model_runner.py` and was not ported into the adapter.

**Fix.** Put calendar bootstrap in the shared core (Option A) and call it at the start of every turn. Bootstrap once per process, cache the id (module-level or adapter-instance state), and re-create if a `GET /calendars/{id}` shows the store was reset. Pass `calendar_id` into `build_user_message` and emit the `- calendar_id = "{calendar_id}"` line in the `FOR THIS EMAIL` preamble exactly as `model_runner.py:229` does. Decide explicitly whether bootstrap belongs to the engine (one calendar per run, created once in `run_simulation`) or the runner (current model_runner behavior). Engine-owned is cleaner because it is harness-agnostic; if you move it to the engine, the adapter just needs the id handed to it.

**Files.** new shared core, `harness.py`, possibly `engine.py` (if engine owns bootstrap).

**Acceptance criteria.**
- A default run creates at least one calendar and at least one event in the store.
- A scenario with a `CC` criterion can score > 0 on the adapter path.
- The model receives a concrete `calendar_id` in the user message for every turn.

**Tests.** Add a test that runs one `CC` scenario through the adapter with a stub or recorded `claude` output and asserts an event lands in the store with the right `scenario_id`. At minimum, add a unit test asserting `build_user_message` output contains a `calendar_id =` line.

**Depends on.** Section 1 decision. Blocks nothing else but is the headline.

---

### FIX-2: Resolve criteria tokens before grading (date matching is dead)  `[CRITICAL]`

**Problem.** Criteria like `CC-{date}` are never resolved, so the grader's date check is skipped entirely and `CC` collapses to a count-only check. A correct-date event and a wrong-date event score identically.

**Evidence.**
- `engine.py:175-180` `apply_date_substitutions` resolves `email.subject` and `email.body` only. It never touches `success_criteria`.
- Per-scenario grading passes the raw loader `Scenario` with unresolved criteria: `engine.py:407-409`.
- The grader returns `None` for any still-tokenized date and silently skips the check: `grader.py:64-77` and `grader.py:151-154`.

**Root cause.** Anthony's Sprint 5 task 1 ("apply `resolve_tokens` to each criterion at the sim_date its email was served") was never implemented, for either grading path.

**Fix.** Resolve `success_criteria` tokens at the sim_date the email was served.
- Per-email path: extend `apply_date_substitutions` (or add a sibling) so it also resolves each entry in `success_criteria` with the same `resolve_tokens(text, sim_date)` used for subject/body. This is the natural place because `sim_date` is in scope at `engine.py:355`.
- Per-scenario path: criteria belong to individual emails, each served on a possibly different `sim_date`. You must resolve each criterion at the date *its email* was served, not at completion day. Capture each email's served `sim_date` during the run (the controller already records `sim_date` in `delivery_log`, see `flow_controller.py:219-226`) and resolve criteria against that. Build a resolved `Scenario` (or resolved criteria list) before calling `define_grading_system` at `engine.py:409`.

**Files.** `engine.py` (both grading call sites and `apply_date_substitutions`), possibly a small helper in `grader.py`.

**Acceptance criteria.**
- For a `CC-{date}` criterion, an event on the correct date passes and an event on the wrong date fails.
- Per-email and per-scenario grading both apply resolution.
- `grader._extract_date_token` receives a resolved date string (no `{...}`) for date criteria.

**Tests.** Add a grader test: same event, two criteria (`CC-{date}` resolved to the event's date vs resolved to a different date), assert pass and fail respectively. Add an engine test that the criteria handed to the grader contain no `{` after resolution.

**Depends on.** Nothing. Can run in parallel with FIX-1 (different files, but both touch `engine.py`, so coordinate, see Section 5).

---

### FIX-3: Restore token / tool / compaction logging on the live path  `[CRITICAL for the benchmark's stated goal]`

**Problem.** The default path writes no `token_usage.jsonl`, no `tool_calls.jsonl`, and never detects compaction. Sprint 5 explicitly required baselining token logs before and after the migration, and "prove a model goes beyond the context window" depends on compaction detection.

**Evidence.**
- `harness.py:130-174` `run_turn` only calls `_extract_session_id` and discards everything else.
- The full parser with logging lives only in `model_runner.py:243-327` (`_parse_stream_output`), `:155-193` (`_record_usage`), `:196-218` (`_log_tool_call`), with compaction detection at `:273-275` and `:295-297`.
- `GETTING_STARTED.md:111-115` admits the adapter path does not trigger this.

**Root cause.** The stream-json parser and its logging side effects were not ported into the adapter.

**Fix.** Move `_parse_stream_output`, `_record_usage`, `_log_tool_call`, and the `_stats` accounting into the shared core, then call the parser from `ClaudeCodeAdapter.run_turn` instead of the bare `_extract_session_id`. Keep the message-id dedup (`model_runner.py:255-282`), it is correct and prevents double-counting. Preserve the `TOKEN_LOG_PATH` / `TOOL_LOG_PATH` env knobs and the `atexit` summary.

**Files.** new shared core, `harness.py`, `bench_logger.py` (already used by the logging functions, no change expected).

**Acceptance criteria.**
- A default run appends rows to `token_usage.jsonl` and `tool_calls.jsonl`.
- Rows split `cache_creation_input_tokens` vs `cache_read_input_tokens` (proves caching is live).
- `COMPACTION fired` prints if a compaction event appears in the stream.
- No double-counted usage rows (dedup intact).

**Tests.** Feed a recorded multi-message stream-json sample (one with a duplicate `message.id`) into the parser and assert usage is counted once and tool calls are logged. Save the sample as a fixture under `tests/`.

**Depends on.** Section 1 decision. Independent of FIX-1/FIX-2 in logic but shares files with FIX-1 (the adapter and core), so sequence within the runner track.

---

## 3. Priority 1: correctness and comparability

### FIX-4: One system prompt, used everywhere  `[HIGH]`

**Problem.** The adapter's `_DEFAULT_SYSTEM_PROMPT` (`harness.py:238-251`, ~13 lines) is missing the delete-only rule, list-call rationing, event-before-todo ORDER, context-reuse guidance, reply format, and the "end immediately" directive that `model_runner.py:97-121` (~25 lines) has. The default path runs a weaker prompt, so the model over-acts and wastes rounds. Two runs are not comparable.

**Fix.** Define the system prompt once in the shared core and have both the adapter and (if retained) the runner reference it. Use the longer `model_runner.py` version as the canonical text. Keep injecting it via `--append-system-prompt` on email 1 only (correct, since `--resume` carries it forward).

**Files.** shared core, `harness.py`.

**Acceptance.** Adapter and runner emit byte-identical system prompts; a grep for prompt text finds exactly one definition.

**Depends on.** Section 1. Same files as FIX-1/FIX-3.

---

### FIX-5: Decide and implement free-text criteria handling  `[HIGH, needs a product decision]`

**Problem (bigger than "free-text auto-passes").** Any criterion without a `TC` / `CC` / `RS` / "No action" prefix auto-passes (`grader.py:161-165`). Measured against the real dataset (full breakdown in Appendix C): of 225 sub-criteria, 99 are prefixed and 96 are "no action", but **30 auto-pass as "free-text", and most of those are a splitter bug, not real free-text.** `grader._parse_sub_criteria` splits on commas and `&&` only, which fragments real criteria so their prefix is no longer at position 0: it produces 13 phantom `^` entries, plus `or CC-{date+1- 3PM}` and `Update:  TC-{nextweek-wednesday}` pieces that are genuine `CC`/`TC` checks silently turned into passes. So FIX-5 is two problems: (1) the splitter mangles real prefixed criteria, and (2) genuinely unprefixed criteria ("Delegate Task", "Add task", "Find when quarterly earnings call is") auto-pass and inflate `max_score`.

**Decision required (pick one, document it in `docs/GRADER.md`):**
1. **Explicit "ungraded" marker.** Free-text criteria are excluded from `max_score` and reported in a separate `ungraded` list. Honest and low-risk. Recommended as the default.
2. **Rubric / keyword rule.** Map known free-text intents ("Delegate", "Flag", "Add task") to checkable conditions (for example "Delegate" requires a `send_email`, "Add task" requires a `TC`). More work, more signal.
3. **Strict fail-closed.** Unknown criteria fail unless matched. Highest signal, highest risk of false negatives. Only if the dataset's free-text strings are normalized first.

**Fix.** Implement the chosen rule in `grader._check_criteria` and adjust `define_grading_system` so `max_score` reflects the decision (option 1 means free-text does not count toward `max_score`).

**Files.** `grader.py`, `docs/GRADER.md`, possibly `engine.py` (results shape if you add an `ungraded` field).

**Acceptance.** Free-text criteria no longer silently pass-and-count. Behavior is documented. Existing tests updated.

**Depends on.** A human decision (use `AskUserQuestion` equivalent before coding). Logic is in `grader.py`, parallel to the runner track.

---

### FIX-6: Clean failure handling on harness crash or timeout  `[HIGH]`

**Problem.** When `adapter.run_turn` raises (timeout or non-zero exit), the engine logs, calls `mark_served`, and `continue`. That marks the email as served with no failure flag, skips grading entirely (implicit silent 0 with no record), and leaves the session id un-updated so the next email in the chain resumes a stale or missing session.

**Evidence.** `engine.py:364-376` (catch, `mark_served`, `continue`). `flow_controller.py:219-226` (delivery_log entry has no error field). `harness.py:161-173` (`session_id` is only stored after a successful run; on exception it is never updated).

**Fix.**
- Record the failure: add an `error` / `status` field to the delivery_log entry (`flow_controller.mark_served` or a new method) so a crashed turn is distinguishable from a clean one.
- Grade it as an attempted-but-failed 0 rather than skipping, OR explicitly tag it "not attempted" in results. Decide which, but do not let it vanish.
- Reset session state for that scenario on failure (clear or re-init `_sessions[scenario_id]`) so the next email starts a fresh session instead of resuming a dead one. Consider an optional one-shot retry for transient timeouts (config-gated, off by default).

**Files.** `engine.py`, `flow_controller.py`, `harness.py`.

**Acceptance.** A forced crash (point the subprocess at a bad flag in a test) yields a delivery_log row marked failed, a graded 0 or explicit not-attempted, and a clean session for the next email. The chain does not break.

**Depends on.** Independent logic, touches `engine.py` (coordinate with FIX-1/FIX-2 on that file).

---

### FIX-7: Single source of truth for MCP config, hidden tools, prompt  `[MEDIUM, free if Option A]`

**Problem.** `_MCP_CONFIG`, `HIDDEN_TOOLS`, `_DISALLOWED_TOOLS` are defined identically in both `harness.py:60-74` and `model_runner.py:49-67`. Update one and they silently diverge, and tool filtering controls which tools the model can call (a core signal).

**Fix.** Falls out of Option A automatically: define these once in the shared core and import. If you took Option B, extract them into a tiny `config.py` both files import.

**Acceptance.** Exactly one definition of each constant; both paths import it.

**Depends on.** Section 1.

---

### FIX-8: Make continuity and CLI knobs reach the live path  `[HIGH]`

**Problem.** `CONVERSATION_CONTINUITY` is read only by `model_runner.py:47`. The engine never passes `conversation_continuity` to `get_adapter` (`engine.py:325-331`), so on the live adapter path it is hard-wired `True` with no way to A/B fresh-per-email. Likewise the CLI args (`--model`, `--reasoning`, `--api-base`, `--openrouter`) reach the adapter but not the `model_runner` fallback (`engine.py:378` calls `run_model_turn` with three args; the runner reads `CLAUDE_MODEL` / `CLAUDE_REASONING` from env, so `CLAUDE_MODEL=...` is silently ignored on the live path and `MIGRATION_AI_LANE.md:27` is wrong about it).

**Fix.** With Option A there is one path, so wire `CONVERSATION_CONTINUITY` and all CLI knobs into it from one place (read the env in `engine.run_simulation` or in the core, and thread CLI args through `get_adapter`). If you keep a fallback runner, give it the same parameter interface so the two are interchangeable. Update `GETTING_STARTED.md` so the env table states which knob applies.

**Files.** `engine.py`, shared core / `harness.py`, `GETTING_STARTED.md`, `MIGRATION_AI_LANE.md`.

**Acceptance.** `CONVERSATION_CONTINUITY=0 python engine.py` actually runs fresh sessions on the live path. `--model claude-sonnet-4-6` runs Sonnet (verify in a turn log). The docs match the code.

**Depends on.** Section 1.

---

## 4. Priority 2: telemetry, reachability, docs, polish

### FIX-9: Compaction and token reporting as a grading dimension  `[HIGH, Anthony task 4]`

**Problem.** The score sheet cannot tell "succeeded within the context window" from "succeeded through compaction," and context-window-exceeded scenarios are not reported separately. `grader.define_grading_system` returns only `score / max_score / details` (`grader.py:207-217`); `run_simulation` returns no token or compaction fields (`engine.py:453-463`).

**Fix.** Track per-scenario compaction (the parser from FIX-3 already detects events, attribute them to `scenario_id`). Add `compaction_triggered`, `context_window_exceeded`, and token totals to the per-scenario result and to the aggregate returned by `run_simulation`. In the printed breakdown, split scenarios into "within window" vs "through compaction."

**Files.** shared core (emit per-scenario compaction), `engine.py` (thread it into results and the printed summary), `bench_logger.py` (new summary section).

**Acceptance.** Results dict has the new fields; the summary shows the split. Depends on FIX-3 (needs the parser on the live path).

---

### FIX-10: Make compaction actually reachable  `[MEDIUM]`

**Problem.** The longest chain in `Emails.xlsx` is about 5 emails (~15-25K tokens) versus a 200K window, so compaction never fires and Sprint 5's "completes through compaction instead of crashing" is never demonstrated. Acknowledged in `MIGRATION_AI_LANE.md:39`.

**Fix.** Add a synthetic stress fixture: either a 50+ email chain scenario in a separate workbook (so the default dataset stays clean) or a generator that produces one. Run it through the live path and confirm a compaction event appears and the chain still completes and grades. This is the actual proof the migration was for.

**Files.** new fixture (for example `stress_chain.xlsx` or a small generator), a test or a documented manual run.

**Acceptance.** A documented run shows `COMPACTION fired` at least once and the long-chain scenario still produces a score. Depends on FIX-3 and FIX-9 for the signal to be visible.

---

### FIX-11: Documentation truth pass  `[MEDIUM]`

Fix the docs that currently contradict the code (most resolve automatically once Option A lands, but verify):
- `GETTING_STARTED.md:82-115`: env table claims `TOKEN_LOG_PATH` / `TOOL_LOG_PATH` / `CONVERSATION_CONTINUITY` are read by "model_runner.py / the adapter." State which path actually reads each.
- `MIGRATION_AI_LANE.md:27`: `CLAUDE_MODEL=... python engine.py` is wrong on the live path. Fix or remove.
- `MCP.md`: add the explicit statement Sprint 5 asked for, that the MCP tool surface is identical across harnesses and adapters differ only in launch and resume (the code enforces this; the doc never says it).
- `README.md`: the endpoint map omits `GET /calendars/` (the list route the grader relies on, `app/routers/calendar.py:57-61`). Add it.

**Acceptance.** No doc instructs a setup the code does not honor.

---

### FIX-12: Make a bad `scenario_id` visible instead of a silent 0  `[MEDIUM, needs a decision]`

**Problem.** `scenario_id` is a roughly 10-digit md5-derived int (`pipeline.py:36-37`) that the model must copy verbatim into every write. If it garbles it, the write 404s, nothing is created, and the email scores 0 with no signal that the cause was a bad id rather than a bad decision. The old SDK runner force-injected the id; the subprocess cannot.

**Decision: D3 (detect-and-log at the MCP server), pre-defaulted in the Execution Kickoff.** The options were:
- (a) Keep current behavior: bad id, silent 0. Defensible but undiagnosable.
- (b, chosen) MCP server detect-and-log: in `create_todo` / `create_event` / `send_email` (`mcp_server/server.py`), when the downstream API returns 404 for a missing `scenario_id`, emit a distinct stderr line (`bad scenario_id: expected vs seen`) and increment a counter, then surface the error to the model as today. Cheap, harness-agnostic, makes silent 0s visible. Keeps "the model must pass it" as part of the test.
- (c) Runner post-validation: after each turn, verify every created todo/event carries the expected `scenario_id` and fail the email loudly otherwise. More thorough, more coupling.

Implement (b) unless the owner overrides. Document the chosen behavior in `docs/api_reference.md`.

**Fix.** At minimum, log when a write fails scenario-id validation so silent 0s become diagnosable. Optionally, reconsider using a short human-typable id at the MCP boundary.

**Files.** `mcp_server/server.py` (or `app/routers/*` for detection), `model_runner` / core for post-validation.

**Acceptance.** A garbled-id run leaves a clear log line. Behavior documented in `docs/api_reference.md`.

---

### FIX-13: Tighten response models  `[LOW]`

`TodoResponse` (`app/models/todo.py:29`), `EventResponse` (`app/models/calendar.py:21`), and the `Email` response model (`app/models/email.py:23`) declare `scenario_id: Optional[int] = None`, but every create path requires and always sets it. Make these non-optional in responses, or document why they are optional. Cosmetic, do last.

---

## 5. Ultracode orchestration plan

The fixes share three files heavily (`engine.py`, `harness.py`, `grader.py`), so naive parallel worktrees will conflict. Sequence around the shared seam.

**Phase 0 (sequential, blocking): the decision and the scaffold.**
- Resolve Section 1 (Option A vs B). If A, one agent creates the shared core module and moves the canonical constants, system prompt, `build_user_message`, `bootstrap_calendar`, and the stream-json parser into it, then points `ClaudeCodeAdapter` at it. This single refactor is the foundation for FIX-1, FIX-3, FIX-4, FIX-7. Do not parallelize this. Verify the engine still runs end-to-end (even against a stubbed `claude`) before fanning out.
- Resolve the two product decisions up front with the user: FIX-5 (free-text rule) and FIX-12 (scenario_id safety net). Use a question tool, do not guess.

**Phase 1 (parallel, three file-disjoint tracks).**
- **Track A, runner/core:** FIX-1, FIX-3, FIX-4, FIX-7, FIX-8. All in the shared core plus `harness.py`. One agent or a short sequential chain, since they touch the same files.
- **Track B, grading:** FIX-2, FIX-5, FIX-9. All in `grader.py` plus the grading call sites and results shape in `engine.py`. FIX-9 depends on FIX-3 emitting per-scenario compaction, so either land FIX-3 first or have Track B consume a stub interface and integrate at Phase 2.
- **Track C, independent:** FIX-10 (fixture), FIX-11 (docs), FIX-13 (models). No conflicts with A or B.
- **The conflict point is `engine.py`.** FIX-2, FIX-6, FIX-8, FIX-9 all edit it. Assign all `engine.py` edits to a single integrating agent, or land them sequentially and rebase, rather than parallel worktrees on `engine.py`.

**Phase 2 (sequential): integration and adversarial verification.**
- Merge tracks, resolve `engine.py` seams.
- Run the full test suite (`python -m pytest tests/ -v`). It must stay green; update tests that encoded the old broken behavior, and add the new tests each fix names.
- Do a real end-to-end run: start uvicorn, `python engine.py Emails.xlsx`, and confirm the global definition of done below. Then run the FIX-10 stress chain and confirm compaction fires.
- Adversarial gate: for each P0 fix, a fresh agent tries to prove it is still broken (create an event on the adapter path, grade a wrong-date event, find an empty `token_usage.jsonl`). Only close a fix when the skeptic fails to break it.

**Verification commands.**
```bash
source venv/bin/activate
python -m uvicorn app.main:app --reload      # terminal A
python engine.py Emails.xlsx                  # terminal B, default = adapter path
python -m pytest tests/ -v
ls -la token_usage.jsonl tool_calls.jsonl     # must be non-empty after a run
claude --help                                 # flags already verified on 2.1.158
```

---

## 6. Global definition of done

The remediation is complete when a default `python engine.py Emails.xlsx` run on the adapter path satisfies all of:

1. At least one calendar and one event exist in the store after the run (FIX-1).
2. A `CC-{date}` criterion passes only when the event is on the correct date (FIX-2).
3. `token_usage.jsonl` and `tool_calls.jsonl` are non-empty and show cache reads vs creation (FIX-3).
4. The adapter and any retained runner use one identical system prompt (FIX-4).
5. Free-text criteria follow the documented rule and do not silently pass-and-count (FIX-5).
6. A forced harness crash produces a logged failure and a graded result, and the chain continues with a clean session (FIX-6).
7. MCP config, hidden tools, and prompt have exactly one definition each (FIX-7).
8. `CONVERSATION_CONTINUITY=0` and `--model <m>` visibly change behavior on the live path (FIX-8).
9. Results distinguish "succeeded within window" from "succeeded through compaction" (FIX-9).
10. The stress chain demonstrably triggers compaction and still scores (FIX-10).
11. No doc contradicts the code (FIX-11).
12. A bad `scenario_id` is diagnosable, not a silent 0 (FIX-12).
13. Test suite at the realistic target from the Execution Kickoff (0 failed, 0 errors with `uvicorn` running, after the stale-test cleanup), with new tests for FIX-1 through FIX-3 and FIX-5 and FIX-6.

---

## 7. Supporting other harnesses (what it takes)

Today the benchmark only drives `claude -p`. This section is the recipe for changing that. The good news is the split is already in the right place: the *benchmark* is harness-neutral, only the *runner* is Claude-specific. None of this is required to ship the P0/P1 fixes above; it is the follow-on that makes the "any harness" claim true.

### 7.1 What is already portable (no per-harness work)

These need zero changes to add a new harness, because they never inspect the agent, they only read the store:

- The **MCP server** (`mcp_server/server.py`). The single tool surface, a thin HTTP wrapper over the API. Any MCP-capable harness (Claude Code, Codex, Cursor, Cline, a custom Agent-SDK loop) connects to the same 22 tools.
- The **store, grader, flow controller, loader** and the **engine contract** (`run_turn(email, sim_date, scenario_id)` then block then diff-grade).
- The **`HarnessAdapter` interface** (`harness.py:77-97`). This is the correct seam. Adding a harness means adding one subclass.

### 7.2 What is Claude-Code-specific (must be reimplemented per harness)

All of this lives in `ClaudeCodeAdapter` and the shared core (after Option A):

- The subprocess command and its flags: `-p`, `--output-format stream-json`, `--mcp-config`, `--strict-mcp-config`, `--disallowed-tools`, `--append-system-prompt`, `--resume`, `--effort`, `--permission-mode bypassPermissions`. None of these exist on other CLIs.
- **Session continuity** via `--resume <session_id>`, where the id is scraped from Claude's stream-json `system/init` event (`harness.py:219-235`). Other harnesses have different or no resume mechanism.
- **Compaction** detection, which parses Claude's stream-json compaction events (`model_runner.py:273-275, 295-297`). Compaction is a Claude Code feature.
- **Telemetry** parsing (tokens, tool calls), which reads Claude's stream-json schema.

### 7.3 The per-harness adapter recipe

To add harness X, write a `HarnessAdapter` subclass that does each of the following. Items 1 to 3 are mandatory; 4 to 6 are needed for full parity.

1. **Point X at the MCP server.** Reuse the same MCP config; only the config *format* differs per harness. The tool surface must stay identical (that is the fairness invariant in 7.4).
2. **Run one turn non-interactively and block.** One email in, control returns only after X is done acting. Whatever X's "headless / print / batch" mode is.
3. **Inject the prompt.** The shared core must expose two shapes, because harnesses split it differently:
   - CLI subprocess harnesses (claude, codex) take system + user concatenated into the prompt argument. Provide a `build_prompt(...)`.
   - In-process SDK harnesses (e.g. the Claude Agent SDK, or an OpenAI Agents loop) take the system prompt as a separate option and the email as the user message. Provide a `build_user_message(...)`.
   This is exactly the `build_prompt` vs `build_user_message` split prototyped on the `architecture-update-with-agents-sdk` branch (`harness/base.py`). Not every harness is a subprocess; the interface allows in-process too.
4. **Session continuity.** Map `scenario_id` to X's session handle and resume it for emails 2..N. If X cannot resume sessions, you have two honest choices: (a) accept that the continuity and compaction test does not apply to X and report it as "not supported," or (b) manually re-send the accumulated conversation each turn. Do not pretend continuity happened if it did not.
5. **Telemetry and compaction.** Parse X's output for tokens, tool calls, and whatever X does on context overflow (compact, truncate, error). This is what makes FIX-3 and FIX-9 produce numbers for X. If X emits nothing parseable, log that those dimensions are unavailable for X rather than emitting zeros.
6. **Non-interactive tool execution.** X must run tools without a human approving each one. Claude uses `--permission-mode bypassPermissions`; every harness has its own auto-approve switch. Find X's.

Then register X in `HARNESS_REGISTRY` (`harness.py:254-257`) and it is selectable via `--harness X`.

### 7.4 Fairness invariants (so a cross-harness score means something)

For two harness runs to be comparable, these MUST be identical across harnesses:

- the MCP tool surface (already enforced by the shared `_MCP_CONFIG` / hidden-tools list),
- the scenarios and emails, and the `scenario_id` contract,
- the grader,
- the secretary system-prompt guidance (the same text, injected via each harness's own mechanism).

These are ALLOWED to vary, and the variance is precisely what the benchmark measures:

- how the harness manages context (compaction vs truncation vs error vs manual re-send),
- the harness's internal agent loop and round count,
- session-resume mechanics.

### 7.5 Recommended change that makes every future harness cheaper

Tool hiding today is done with Claude's `--disallowed-tools` flag, which is per-harness and not portable. **Move the hiding into the MCP server instead:** gate the admin tools (`create_scenario`, `delete_scenario`, `create_calendar`, etc.) behind a `BENCH_MODE` env var so the server simply does not register them during a benchmark run. Then every harness automatically gets the trimmed surface with zero per-harness flag work, and the fairness invariant in 7.4 (identical tool surface) is enforced in one place instead of N. This also kills FIX-7's duplication at the source. Strongly recommended before adding harness #2.

### 7.6 Optional FIX-14: prove the abstraction with one real second harness  `[only if multi-harness is a near-term goal]`

The Sprint 5 "done when" for the abstraction was "switching harness is one CLI/env value with zero changes to engine, grader, or MCP." That is satisfied by the *interface*, but it is unproven until a second harness actually runs. To prove it: implement one real adapter (Codex CLI is the natural candidate, it is already MCP-capable and stubbed at `harness.py:182-200`), apply 7.5 so tool hiding is server-side, and confirm a full run produces a score with no edits to `engine.py`, `grader.py`, or the MCP server. Acceptance: two harnesses run the same scenario set and the only difference in the codebase between them is the adapter class plus launch config. Decide the target harness before starting.

---

## Appendix A: verified findings (severity after adversarial review)

| ID | Finding | Severity | Sprint owner |
|----|---------|----------|--------------|
| FIX-1 | Adapter path cannot create calendar events (no bootstrap, no `calendar_id`, `create_calendar` disallowed) | CRITICAL | Eyasu (adapter) / Miguel (runner) |
| FIX-2 | `success_criteria` tokens never resolved, grader date matching is dead | CRITICAL | Anthony (task 1) |
| FIX-3 | No token/tool/compaction logging on the live path | CRITICAL (for stated goal) | Eyasu / Miguel |
| FIX-4 | Adapter system prompt is ~half the runner's, missing key guidance | HIGH | Eyasu |
| FIX-5 | Free-text criteria auto-pass and inflate max_score | HIGH | Anthony (task 3) |
| FIX-6 | Crash/timeout marks email served, skips grading, leaves stale session | HIGH | Nikita (task 2) |
| FIX-7 | MCP config / hidden tools / prompt duplicated in two files | MEDIUM | Eyasu / Miguel |
| FIX-8 | `CONVERSATION_CONTINUITY` and CLI knobs do not reach the live path | HIGH | Nikita / Miguel |
| FIX-9 | No compaction/token dimension in grading output | HIGH | Anthony (task 4) |
| FIX-10 | Compaction never reachable (max 5-email chains) | MEDIUM | Miguel (done-when) |
| FIX-11 | Docs contradict code (env table, CLAUDE_MODEL, MCP.md, calendar list route) | MEDIUM | all |
| FIX-12 | Garbled `scenario_id` is a silent 0 | MEDIUM | Anthony (task 2 follow-through) |
| FIX-13 | Response models mark `scenario_id` optional though always set | LOW | api |

## Appendix B: do not chase these (adversarially refuted)

- The stream-json message-id dedup is **correct** and should be preserved, not rewritten. It is missing from the adapter only because the adapter does no logging at all (that is FIX-3, not a dedup bug).
- `resume_session` being a no-op in `ClaudeCodeAdapter` is **intentional**, resume is handled inside `run_turn` via `--resume`. Do not "implement" it.
- Day-100 overflow is **prevented** by offset clamping (`flow_controller.py:152-155`); a test asserts `remaining_active <= 5`. The "no leftover" invariant essentially holds. The only ungraded tail is a handful of single-email scenarios that activate on the last days, which is acceptable and reported as `remaining_active`.
- The `CodexAdapter` docstring ("MCP config identical to ClaudeCodeAdapter") is forward-looking design documentation, not a false claim.
- Every `claude` CLI flag the code uses (`--tools ""`, `--effort`, `--resume`, `--disallowed-tools`, `--append-system-prompt`, `--strict-mcp-config`, `--mcp-config`) exists and behaves as intended on the installed `claude` 2.1.158. The flag surface is not a risk on this machine.

## Appendix C: free-text criteria data (for FIX-5)

Measured from `Emails.xlsx` via `loader.load_scenarios` + `grader._parse_sub_criteria`. This is the real data FIX-5's decision rests on.

- 109 scenarios, **225 sub-criteria total**: 99 prefixed (`TC`/`CC`/`RS`), 96 "no action", **30 auto-pass as free-text** (across 15 distinct strings).
- The 30 fall into three buckets, and only the third is genuinely "free-text":

**Bucket 1, splitter artifacts (NOT free-text, real checks broken by the parser):**
| count | string | what it should be |
|---|---|---|
| 13 | `^` | a phantom token; the cell uses `^` somewhere the splitter leaves stranded. Investigate the source cells, strip/ignore `^`. |
| 2 | `or  CC-{date+1- 3PM}` | the second half of a "`CC-{x}` or `CC-{y}`" alternative; prefix no longer at position 0 |
| 2 | `or CC-{date+2- 11AM}` | same, an "or" alternative |
| 1 | `Update:  TC-{nextweek-wednesday}` | a real `TC` check prefixed with "Update: " |

These prove the splitter mangles real `CC`/`TC` criteria into auto-passes. Fix `_parse_sub_criteria` to handle `or` alternatives, leading labels ("Update:"), and `^`.

**Bucket 2, natural-language action criteria (no prefix, describe a real action):**
| count | string |
|---|---|
| 1 | `delete meeting {date-14th} {date-11AM}` |
| 1 | `Remove meeting on {date-1:14PM}` |
| 1 | `create new meeting on {date-3PM}` |

These describe real expected events but use no `TC`/`CC`/`RS` prefix, so they auto-pass. Either normalize them in the dataset (rewrite as `CC-`/`RS-`) or have the grader recognize the verbs.

**Bucket 3, genuinely free-text intents (the actual FIX-5 question):**
| count | string |
|---|---|
| 2 | `Add task` |
| 1 | `Flag` |
| 1 | `insufficient info` |
| 1 | `Delegate Task` |
| 1 | `Create to:do` |
| 1 | `Add task to do list` |
| 1 | `Find when quarterly earnings call is` |
| 1 | `Must not re:add` |

`Flag` + `insufficient info` is one criterion ("Flag, insufficient info") split on its comma. The D2 default ("ungraded marker") applies to this bucket: exclude from `max_score`, report separately. Buckets 1 and 2 are bugs to fix, not free-text to mark.
