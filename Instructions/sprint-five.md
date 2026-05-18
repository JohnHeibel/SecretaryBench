
Right now the AI lane (`model_runner.py`) talks to the model with the **Anthropic SDK directly**. It hand-rolls its own agent loop, its own message chain, and its own "context management" — which today just means it keeps the per-scenario conversation in memory and throws it away when the scenario finishes. There is no compaction. When a scenario chain's conversation grows past the model's context window, the API rejects it and the turn crashes.

- We are **required to use context management tools**.
- We need to **prove a model that goes beyond a model's context window**.
- We need **automatic compaction**.

The current SDK design cannot do any of that. A harness (Claude Code, Codex) gives us automatic compaction and context management for free. Our MCP server is already harness-agnostic — MCP.md already documents Claude Code, Cursor, and Codex all connecting to it. So the migration is contained: we swap the SDK call path inside `model_runner.py` for a harness, keep the MCP server as the single tool surface, and keep the engine contract the same.

This is what we are moving to:

1. Migrate from the SDK to the Claude Code harness.
2. Migrate the model toggle to a CLI parameter passed on the simulation run (optional parameter and default model [can be claude haiku]).
3. Design it so any harness can be dropped in through the MCP, and new models from those harnesses are supported.
4. Allow Claude Code through OpenRouter.
5. Allow the model's reasoning flag to be passed through.

**On the token usage question** — "are we using too many tokens, is that normal, or do we optimize?" The answer is both, and we have to separate them or the number is meaningless. Long chains, the store accumulating over 100 days, and context exceeding the window is the *thing we are testing* — that token cost is the benchmark, do not optimize it away. Tokens wasted on an uncached system prompt, verbose tool schemas, redundant `list_*` calls, and re-read resources is *noise* — optimize that so it doesn't pollute the signal. The SDK runner already fought the noise (schema minification, terse tool descriptions, hidden tools, a byte-stable system block for cache hits); the migration must not lose those wins. Claude Code's automatic prompt caching actually helps us here. We report tokens as a cost-to-complete metric, not something to globally minimize. Baseline the existing `token_usage.jsonl` and `tool_calls.jsonl` before migrating so the after-numbers are comparable.

Here are the tasks:

**Miguel** — SDK → Claude Code harness migration (you own the AI lane, you own the swap)

1) Spike: drive Claude Code non-interactively against the existing MCP server. Block until it finishes one email. Verify the actual flag surface against the installed `claude` version — don't assume, confirm (`--print`, `--mcp-config`, `--model`, `--append-system-prompt`, `--output-format`, the session flags, permission mode).
2) Preserve the engine contract. `run_model_turn(email, sim_date, scenario_id) -> None` must still block until the harness finishes that one email, so the engine's before/after diff grading keeps working untouched.
3) Per-scenario session continuity — this is the context-window test. Map `scenario_id` to a harness session and resume it for emails 2..N in a chain so the conversation accumulates and eventually triggers the harness's automatic compaction. This replaces `_scenario_messages` and `scenario_completed()`.
4) Re-apply the token-noise mitigations in harness form: system prompt via the append-system-prompt path, trimmed tool surface, confirm prompt caching is active.
5) Reasoning flag: expose extended-thinking / reasoning effort as a passthrough.

Done when: a scenario chain that exceeds the context window completes through compaction instead of crashing, and per-email diff grading still produces scores.

**Eyasu** — Harness abstraction layer (the "drop any harness" design)

1) Define a `HarnessAdapter` interface: `start_session`, `run_turn`, `resume_session`, `end_session`. `model_runner` becomes one consumer of it.
2) Implement the Claude Code adapter (with Miguel's spike). Stub the Codex adapter so the second harness is a config change, not a rewrite.
3) Keep MCP as the single tool surface — adapters differ only in how they launch and resume the harness process, never in tools. Document that boundary in MCP.md.
4) OpenRouter path for Claude Code: wire the base-URL / auth-token env (verify OpenRouter's Anthropic-compatible endpoint) so non-Anthropic models run through the same adapter.

Done when: switching harness or model is one CLI/env value with zero changes to engine, grader, or MCP.

**Nikita** — CLI model toggle + pooling/engine under the harness (model and pooling)

1) `engine.py` argparse: `--model`, `--harness`, `--reasoning`, `--api-base`/`--openrouter`, threaded through `run_simulation` into the adapter and down to the harness flags/env. Today `__main__` only takes a path.
2) Validate the pool/chain lifecycle survives the swap: per-email diff attribution still correct when the model call is a subprocess; session lifecycle keyed to pool transitions (start on activate, resume per email, end on completion — this is the hook Miguel's session map needs); harness timeout or crash rolls back cleanly without breaking the chain.
3) Confirm "no leftover emails after 100 days" still holds, and that long-chain scenarios actually reach compaction in a real run.

Done when: `python engine.py Emails.xlsx --harness claude-code --model <m> --reasoning high` runs 100 days, pools drain, scores produce.

**Anthony** — Grading under the new model path

1) Resolve criteria tokens before grading. Apply `engine.resolve_tokens` to each criterion at the sim_date its email was served, so `CC-{date}` and content+date checks actually fire at scenario level. This closes the gap GRADER.md already flags.
2) Own the `scenario_id` contract decision — this is a migration blocker. Today `_inject_keys` force-injects the scenario_id because small models garble it; a harness subprocess can't do that mid-loop. Decide with Miguel and Eyasu: the MCP server derives/defaults it from session context, or the model must pass it correctly and that becomes part of what we test. Grading correctness depends on this — drive it.
3) Decide free-text / unprefixed criteria handling (currently a lenient pass-through). Stricter rule, rubric, or an explicit "ungraded" marker. This is tied to the "what does the final success criteria look like" question already in GRADER.md.
4) Add token/compaction reporting as a grading dimension per the framing above: report context-window-exceeded scenarios separately, and reward correct behavior despite compaction.

Done when: grading is correct against harness-produced state and the score sheet distinguishes "succeeded within window" from "succeeded through compaction."

**Sequencing.** Miguel's spike unblocks everything — start there. Eyasu's interface and Miguel's runner co-develop against each other. Nikita's CLI can start in parallel against the interface. Anthony's items 1, 3, and 4 are independent and can start now; item 2 is the one cross-cutting blocker and must be resolved jointly between Anthony, Miguel, and Eyasu before the migration merges.

*Yes this was AI-generated.*