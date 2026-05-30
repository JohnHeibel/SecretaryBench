# Sprint 5 Remediation Plan

**Branch:** `sprint5-remediation`
**Status:** plan only, no code changed yet
**Audience:** an ultracode session (multi-agent workflow) or a human implementing the fixes
**Source of findings:** a 46-agent adversarial audit of `main` plus first-hand reads of every core file. Every claim below was verified against the source, not just the docs.

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

**Required decision before P0 work starts.** Pick one:

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

**Problem.** Any criterion without a `TC` / `CC` / `RS` / "No action" prefix auto-passes (`grader.py:161-165`). There are 22 such free-text criteria in `Emails.xlsx` ("Delegate Task", "Flag, insufficient info", and similar). A scenario mixing prefixed and free-text criteria can be passed by satisfying only the prefixed subset, and free-text criteria inflate `max_score` while testing nothing.

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

**Decision required:** keep "the model must pass it, and that is part of the test" (current, defensible), or add a safety net. If a safety net: the MCP server could detect a write whose `scenario_id` does not exist and log it distinctly, or the runner could post-validate that the model used the expected id. Do not silently swallow it.

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
13. Full test suite green, with new tests for FIX-1 through FIX-3 and FIX-5 and FIX-6.

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
