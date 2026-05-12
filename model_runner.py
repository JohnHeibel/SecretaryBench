"""
model_runner.py — AI lane for SecretaryBench.

Drives an LLM through the MCP server (stdio subprocess) once per email.
The model sees a curated subset of MCP tools (admin/destructive ones are
hidden — see HIDDEN_TOOLS).

Architecture:

    engine.py (sync)
        │  run_model_turn(email, sim_date, scenario_id)
        ▼
    model_runner.py
        │  - Anthropic SDK (model layer; one ThreadPoolExecutor for deadline-
        │    enforced API calls)
        │  - MCP stdio client (tool layer, harness-agnostic)
        ▼
    mcp_server (subprocess, stdio)
        │  HTTP
        ▼
    FastAPI store

Key decisions (see modelRunner.md for the why):

- One MCP session per benchmark run. Lazy-init on first call to
  run_model_turn, torn down at process exit.
- A dedicated asyncio event loop runs in a daemon thread so the sync engine
  can submit work via run_coroutine_threadsafe.
- One shared calendar bootstrapped via the MCP `create_calendar` tool. The
  bootstrap re-verifies on every turn so a uvicorn restart between runs
  self-heals (D2).
- Tool surface is FILTERED: admin tools (create_scenario, delete_*,
  add_scenario_email) and create_calendar/health_check are hidden from the
  model so it can't tamper with the benchmark or spawn extra calendars.
- Anthropic API calls run in a ThreadPoolExecutor and are bounded by the
  deadline via Future.result(timeout=...) — so SDK-internal 429 retry sleeps
  cannot push past PER_TURN_TIMEOUT_S (B4). max_retries=0 also short-circuits
  the SDK's silent backoff.
- No fixed cap on tool rounds; 25-round backstop and 300s per-turn deadline
  are the safety bounds. Both crash the run with the scenario_id logged.
- Run summary printed at process exit via atexit.
"""
from __future__ import annotations

import asyncio
import atexit
import concurrent.futures
import json
import os
import sys
import threading
import time
from collections import defaultdict
from datetime import datetime, timezone
from typing import Any

from loader import Email

# --- Module-level state (singleton per process) ----------------------------

_loop: asyncio.AbstractEventLoop | None = None
_loop_thread: threading.Thread | None = None

_session: Any = None                          # mcp.ClientSession when open
_session_ready = threading.Event()            # set once initialize() returns
_session_lock = threading.Lock()              # guards _ensure_session
_shutdown_signal: asyncio.Event | None = None # set on atexit to tear down

_api_executor: concurrent.futures.ThreadPoolExecutor | None = None

_calendar_id: str | None = None
_cached_anthropic_tools: list[dict] | None = None

# Persistent per-scenario conversation. Keyed by int scenario_id. The chain
# accumulates {user, assistant, tool_result} blocks across emails in the same
# scenario so the model retains the context (and tool_use ids) it built earlier.
# Cleared by scenario_completed() once the engine has graded that scenario.
_scenario_messages: dict[int, list[dict]] = defaultdict(list)

# Default ON: per-scenario message chain retained across emails. Set
# CONVERSATION_CONTINUITY=0 to A/B against the legacy fresh-per-turn behavior.
CONVERSATION_CONTINUITY = os.environ.get("CONVERSATION_CONTINUITY", "1") == "1"

_stats: dict[str, Any] = {
    "scenarios_run": 0,
    "tool_calls": defaultdict(int),
    "tool_errors": 0,
    "calendars_created": 0,
    "turn_failures": 0,
    "rounds_total": 0,
    "input_tokens": 0,
    "output_tokens": 0,
}

MODEL_NAME = "claude-haiku-4-5"
PER_TURN_TIMEOUT_S = 300.0
ROUND_BACKSTOP = 25
SESSION_OPEN_TIMEOUT_S = 30.0

# Per-round token usage is appended here as JSONL for offline analysis.
# Override with TOKEN_LOG_PATH env var; set to empty string to disable.
TOKEN_LOG_PATH = os.environ.get("TOKEN_LOG_PATH", "token_usage.jsonl")
_token_log_lock = threading.Lock()

# Per-tool-call audit log; one JSONL row per model tool_use. Lets us answer
# "which scenario/email called list_emails, with what args?" — the per-round
# token log only stores a count of tool_uses, not the names. Empty string
# disables; set TOOL_LOG_PATH=/dev/null to suppress without code change.
TOOL_LOG_PATH = os.environ.get("TOOL_LOG_PATH", "tool_calls.jsonl")
_tool_log_lock = threading.Lock()

# Tools the model should NOT see. The MCP server still exposes them (other
# harnesses keep working); we just don't pass them in the `tools=` list to
# the model.
#
# Design intent: the model sees the same surface a real secretary would have
# — write (create_*), modify (update_*), delete (delete_*), read (list_*,
# get_*), reply (send_email). The benchmark measures whether the model can
# handle reschedules, updates, conflict-avoidance, and duplicate-correction —
# not just creation — so trimming the surface would understate capability.
#
# Hidden categories:
#   - admin: would let the model corrupt the benchmark harness itself.
#   - engine-internal: scenario/email metadata is injected into the user
#     message preamble already, so exposing the read endpoints just gives
#     the model rope to waste a round fetching what it already has.
HIDDEN_TOOLS = frozenset({
    # admin — corrupt the benchmark
    "create_scenario",
    "delete_scenario",
    "add_scenario_email",
    "create_calendar",
    "delete_calendar",
    "health_check",
    # engine-internal — not the model's concern
    "get_email",
    "list_scenarios",
    "get_scenario",
})

# Tools where the runner FORCES the foreign-key field to the correct value
# before forwarding to MCP. Small models (esp. Haiku) garble the 10-digit
# scenario_id hash and the calendar UUID; the benchmark tests "can the model
# decide what action to take", not "can it copy a number". We let the model
# pass whatever it wants — we overwrite before submission.
_FORCE_SCENARIO_ID = frozenset({"create_todo", "create_event", "update_event", "send_email", "list_todos"})
_FORCE_CALENDAR_ID = frozenset({
    "create_event",
    "update_event",
    "delete_event",
    "list_events",
    "get_event",
    "get_calendar",
})


def _record_usage(
    scenario_id: int,
    round_idx: int,
    response: Any,
    messages_in_context: int = 0,
) -> dict[str, int]:
    """Pull usage off a messages.create response, aggregate into _stats, append JSONL row.

    Returns the per-round usage dict so callers can use it for in-turn logging.
    """
    usage = getattr(response, "usage", None)
    row = {
        "input_tokens": getattr(usage, "input_tokens", 0) or 0,
        "output_tokens": getattr(usage, "output_tokens", 0) or 0,
    }
    _stats["input_tokens"] += row["input_tokens"]
    _stats["output_tokens"] += row["output_tokens"]
    _stats["rounds_total"] += 1

    if TOKEN_LOG_PATH:
        tool_uses = sum(1 for b in response.content if b.type == "tool_use")
        line = json.dumps({
            "ts": datetime.now(timezone.utc).isoformat(),
            "scenario_id": scenario_id,
            "round": round_idx,
            "stop_reason": getattr(response, "stop_reason", None),
            "tool_uses": tool_uses,
            "messages_in_context": messages_in_context,
            "continuity": CONVERSATION_CONTINUITY,
            **row,
        })
        try:
            with _token_log_lock, open(TOKEN_LOG_PATH, "a") as f:
                f.write(line + "\n")
        except OSError:
            pass  # logging must never break a run
    return row


def _measure_response(content: str) -> tuple[int, int]:
    """Return (response_bytes, response_items) for a tool response payload.

    response_items counts list items if the payload is a JSON list, or the
    nested events/todos array if it's a dict (get_calendar shape). 0 otherwise.
    Used to detect when list_* responses start growing past the point where
    they cause prompt-bloat or model confusion — see Phase 0 of the scaling
    plan. We do NOT touch the content itself; this is read-only measurement.
    """
    response_bytes = len(content) if isinstance(content, str) else 0
    response_items = 0
    try:
        parsed = json.loads(content)
    except (json.JSONDecodeError, ValueError, TypeError):
        return response_bytes, 0
    if isinstance(parsed, list):
        response_items = len(parsed)
    elif isinstance(parsed, dict):
        # get_calendar returns {"calendar_id": ..., "events": [...]} etc.
        for key in ("events", "todos"):
            value = parsed.get(key)
            if isinstance(value, list):
                response_items = len(value)
                break
    return response_bytes, response_items


def _log_tool_call(
    scenario_id: int,
    round_idx: int,
    tool_name: str,
    raw_args: dict,
    email_index: int,
    is_error: bool,
    response_bytes: int = 0,
    response_items: int = 0,
) -> None:
    """Append one JSONL row per model tool_use.

    raw_args is the model's UNMODIFIED input (pre force-inject) so the log
    reflects what the model actually decided to pass. email_index is 1-based
    within the scenario: 1 = first email (no prior chain context, so
    list_emails is plausibly legit), 2+ = follow-up email (chain already
    holds prior context, so list_emails is likely redundant).

    response_bytes / response_items capture how big the tool's response was
    so we can detect when list_* calls start returning enough data to bloat
    the next round's prompt — the trigger for restoring get_by_id as primary.
    """
    if not TOOL_LOG_PATH:
        return
    line = json.dumps({
        "ts": datetime.now(timezone.utc).isoformat(),
        "scenario_id": scenario_id,
        "round": round_idx,
        "tool_name": tool_name,
        "args": raw_args,
        "email_index": email_index,
        "is_error": is_error,
        "response_bytes": response_bytes,
        "response_items": response_items,
        "continuity": CONVERSATION_CONTINUITY,
    })
    try:
        with _tool_log_lock, open(TOOL_LOG_PATH, "a") as f:
            f.write(line + "\n")
    except OSError:
        pass  # logging must never break a run


def _inject_keys(name: str, args: dict, scenario_id: int, calendar_id: str) -> dict:
    """Force scenario_id and calendar_id to known-correct values."""
    out = dict(args)
    if name in _FORCE_SCENARIO_ID:
        out["scenario_id"] = scenario_id
    if name in _FORCE_CALENDAR_ID:
        out["calendar_id"] = calendar_id
    return out


# --- Async loop in a daemon thread -----------------------------------------

def _ensure_loop() -> asyncio.AbstractEventLoop:
    global _loop, _loop_thread
    if _loop is not None and _loop.is_running():
        return _loop
    _loop = asyncio.new_event_loop()
    _loop_thread = threading.Thread(
        target=_loop.run_forever,
        name="mcp-loop",
        daemon=True,
    )
    _loop_thread.start()
    return _loop


def _run_async(coro, timeout: float | None = None):
    loop = _ensure_loop()
    fut = asyncio.run_coroutine_threadsafe(coro, loop)
    return fut.result(timeout=timeout)


# --- Anthropic API executor (lets us enforce a real wall-clock deadline) ---

def _get_api_executor() -> concurrent.futures.ThreadPoolExecutor:
    global _api_executor
    if _api_executor is None:
        _api_executor = concurrent.futures.ThreadPoolExecutor(
            max_workers=1, thread_name_prefix="anthropic-api"
        )
    return _api_executor


# --- MCP session lifecycle -------------------------------------------------

async def _session_lifecycle() -> None:
    """Long-running coroutine that owns the MCP subprocess + session."""
    global _session, _shutdown_signal, _cached_anthropic_tools
    from mcp import ClientSession, StdioServerParameters
    from mcp.client.stdio import stdio_client

    _shutdown_signal = asyncio.Event()
    params = StdioServerParameters(
        command="bash",
        args=["-c", "uv run python -m mcp_server"],
    )
    try:
        async with stdio_client(params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()
                _session = session
                _session_ready.set()
                await _shutdown_signal.wait()
    finally:
        _session = None
        _cached_anthropic_tools = None  # force re-list on any future reconnect


def _ensure_session() -> None:
    """Idempotent + thread-safe: open MCP session on the daemon loop if not open."""
    if _session is not None:
        return
    with _session_lock:
        if _session is not None:
            return
        loop = _ensure_loop()
        _session_ready.clear()
        asyncio.run_coroutine_threadsafe(_session_lifecycle(), loop)
        if not _session_ready.wait(timeout=SESSION_OPEN_TIMEOUT_S):
            raise RuntimeError(
                f"MCP session did not initialize within {SESSION_OPEN_TIMEOUT_S}s. "
                "Is FastAPI running on localhost:8000? "
                "Try: uv run uvicorn app.main:app --reload"
            )


# --- MCP <-> Anthropic tool bridging ---------------------------------------

async def _list_tools_async() -> list[dict]:
    result = await _session.list_tools()
    return [
        {
            "name": t.name,
            "description": t.description or "",
            "input_schema": t.inputSchema or {"type": "object", "properties": {}},
        }
        for t in result.tools
        if t.name not in HIDDEN_TOOLS
    ]


async def _call_tool_async(name: str, args: dict) -> tuple[str, bool]:
    """Call an MCP tool. Returns (content_string, is_error)."""
    result = await _session.call_tool(name, args)
    parts: list[str] = []
    for block in result.content:
        text = getattr(block, "text", None)
        parts.append(text if text is not None else str(block))
    payload = "\n".join(parts) if parts else "{}"
    is_error = bool(getattr(result, "isError", False))
    if is_error:
        return json.dumps({"error": payload}), True
    return payload, False


# Terse, model-facing descriptions for the 4 visible tools. Replaces the verbose
# MCP-server descriptions (which are written for human/API discovery). The system
# prompt already carries the "how to use these" guidance; the description here is
# just a one-line identity reminder. Override-by-name lets us keep MCP as the
# single source of truth for tool *behavior* while controlling token cost.
_MODEL_TOOL_DESCRIPTIONS = {
    "create_todo": "Create a todo. due_date is ISO 8601 with timezone.",
    "create_event": "Create a calendar event. start/end are ISO 8601 with timezone; end > start.",
    "update_todo": "Modify an existing todo by UUID. Partial — only pass fields you want to change.",
    "update_event": "Modify an existing event by UUID. Full replace — pass every field (title, start, end, scenario_id), not just changed ones.",
    "delete_todo": "Delete a todo by UUID. Use ONLY to remove a duplicate you just created by mistake.",
    "delete_event": "Delete an event by UUID. Use ONLY to remove a duplicate you just created by mistake.",
    "send_email": "Send a reply email.",
    "list_emails": "List emails for the current scenario. ALWAYS pass scenario_id — the unfiltered list grows huge.",
    "list_todos": "List todos for the current scenario. ALWAYS pass scenario_id — the unfiltered list grows huge.",
    "list_events": "List events on the shared calendar.",
    "get_calendar": "Fetch the shared calendar with all its events.",
    "get_todo": "Fetch one todo by UUID.",
    "get_event": "Fetch one event by UUID.",
}


def _minify_schema(schema: dict) -> dict:
    """Strip Pydantic-generated noise (titles, anyOf-null, defaults) from a tool's
    input_schema. Equivalent semantics, fewer tokens sent to the model.

    Pydantic emits `"title": "Calendar Id"` on every property and uses
    `anyOf: [{type:T}, {type:null}]` to express Optional[T]. The model doesn't
    need either — optional-ness is conveyed by the `required` list.
    """
    out: dict = {"type": "object"}
    props: dict = {}
    for name, spec in schema.get("properties", {}).items():
        if "anyOf" in spec:
            non_null = [s for s in spec["anyOf"] if s.get("type") != "null"]
            if len(non_null) == 1:
                spec = non_null[0]
            elif non_null:
                spec = {"anyOf": non_null}
        clean = {k: v for k, v in spec.items() if k not in ("title", "default")}
        props[name] = clean
    out["properties"] = props
    if schema.get("required"):
        out["required"] = schema["required"]
    return out


def _get_anthropic_tools() -> list[dict]:
    """Lazy-memoize the (filtered) tool list. Strips Pydantic schema noise and
    swaps verbose MCP descriptions for terse model-facing ones.
    """
    global _cached_anthropic_tools
    if _cached_anthropic_tools is None:
        raw = _run_async(_list_tools_async(), timeout=30)
        slim = []
        for t in raw:
            slim.append({
                "name": t["name"],
                "description": _MODEL_TOOL_DESCRIPTIONS.get(t["name"], t["description"]),
                "input_schema": _minify_schema(t["input_schema"]),
            })
        _cached_anthropic_tools = slim
    return _cached_anthropic_tools


# --- Calendar bootstrap with self-heal (D2) -------------------------------

def _bootstrap_calendar(sim_date: datetime) -> str:
    """Return a valid shared calendar id, recreating if the store was reset."""
    global _calendar_id

    if _calendar_id is not None:
        # Verify it still exists — uvicorn --reload may have wiped the store.
        try:
            _, is_err = _run_async(
                _call_tool_async("get_calendar", {"calendar_id": _calendar_id}),
                timeout=10,
            )
        except Exception:
            is_err = True
        if not is_err:
            return _calendar_id
        # Stale id — fall through to recreate.
        _calendar_id = None

    aware = sim_date if sim_date.tzinfo else sim_date.replace(tzinfo=timezone.utc)
    iso = aware.isoformat().replace("+00:00", "Z")

    raw, is_err = _run_async(
        _call_tool_async("create_calendar", {"start_date": iso}),
        timeout=30,
    )
    if is_err:
        raise RuntimeError(f"create_calendar failed: {raw}")
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"create_calendar returned non-JSON: {raw!r}") from exc

    if not isinstance(data, dict) or "calendar_id" not in data:
        raise RuntimeError(f"create_calendar response missing calendar_id: {data!r}")

    _calendar_id = data["calendar_id"]
    _stats["calendars_created"] += 1
    return _calendar_id


# --- Prompt construction ---------------------------------------------------

# Static portion of the system prompt — byte-identical across every turn so
# anything that varies per-turn (scenario_id, calendar_id, sim_date) MUST go in
# each user message's preamble, never in this string. Intentionally short: the
# benchmark measures the AI's ability to handle complex requests, so coaching
# the model with worked examples or pattern lookup tables would game the
# capability measurement rather than reveal it.
_STATIC_SYSTEM_PROMPT = """You are an AI executive assistant. For each email, decide the action and call the matching tool — or call no tool. Do not narrate.

TOOLS:
- create_todo: principal must track a task with a deadline.
- create_event: a time-blocked meeting/call is proposed or confirmed.
- send_email: a direct question is asked or a response is explicitly requested.
- update_todo / update_event: an email modifies something you ALREADY created earlier in this scenario. Reschedules use update_event — never delete-then-create.
- delete_todo / delete_event: ONLY to undo a duplicate you just created by mistake. Never to "reschedule" or "cancel" — use update_* for changes; for cancellations the email just doesn't need any action.
- list_emails / list_todos / list_events / get_calendar / get_todo / get_event: ONLY when you need state you don't already remember from this conversation. Prefer NOT to call — you usually remember what you wrote.
- (no tool): FYI / newsletter / auto-confirmation / marketing / status update.

When in doubt, do nothing. Over-acting is the most common failure.

ORDER: if both an event AND a linked todo are needed, create the event first, then pass its returned id as `calendar_event_id` to create_todo.

DATETIMES: ISO 8601 with timezone (e.g. "2000-05-09T14:00:00Z"). Use the simulated date in FOR THIS EMAIL for relative phrases. Defaults: events 09:00, todo due_date 17:00, 1-hour duration.

SCENARIO_ID: copy verbatim from the LATEST FOR THIS EMAIL into every tool call.

CONTEXT: prior emails in this scenario (if any) appear earlier in this conversation along with the tool calls you made for them. Reuse that context — do NOT call list_emails for chain history you already remember.

REPLIES (send_email): recipients = incoming sender. Subject = "Re: <original>" (no double prefix). 1-3 sentences for simple Q's.

END IMMEDIATELY after the tool call(s). Do not emit confirmation text after a tool succeeds — that wastes a round.
"""


def _build_system_blocks() -> list[dict]:
    """Return the static system prompt as a single text block.

    Everything dynamic (scenario_id, calendar_id, sim_date) lives in each user
    message's preamble so the system block stays byte-identical across turns —
    a clean prerequisite for the persistent-conversation path.
    """
    return [{"type": "text", "text": _STATIC_SYSTEM_PROMPT}]


def _build_user_message(
    email: Email, sim_date: datetime, scenario_id: int, calendar_id: str
) -> str:
    """Email body + FOR THIS EMAIL preamble.

    Putting the dynamic context in the user message (instead of system) means a
    persistent conversation can carry prior turns forward: the model sees the
    new FOR THIS EMAIL block on the latest user message and uses those values
    for tool-call arguments without re-reading earlier turns.
    """
    recipients = ", ".join(email.recipients) if email.recipients else "(none)"
    sim_date_str = sim_date.strftime("%B %d, %Y")
    sim_date_iso = sim_date.strftime("%Y-%m-%dT%H:%M:%SZ")
    preamble = (
        "FOR THIS EMAIL:\n"
        f"- scenario_id = {scenario_id}\n"
        f'- calendar_id = "{calendar_id}"\n'
        f"- Today's simulated date is {sim_date_str} (ISO: {sim_date_iso}).\n\n"
    )
    return (
        f"{preamble}"
        f"From: {email.sender}\n"
        f"To: {recipients}\n"
        f"Subject: {email.subject}\n"
        f"Date: {sim_date.strftime('%Y-%m-%d')}\n\n"
        f"{email.body}"
    )


# --- The main entrypoint ---------------------------------------------------

def run_model_turn(email: Email, sim_date: datetime, scenario_id: int = 0) -> None:
    """Send one email to the model and let it act through MCP tools.

    Side-effects (todos, events, replies) land in the FastAPI store via MCP
    tool calls. Returns nothing — the grader reads results back.
    """
    import anthropic  # lazy so a missing package doesn't crash engine.py at import

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise EnvironmentError("ANTHROPIC_API_KEY is not set — cannot run the AI lane.")

    _ensure_session()
    calendar_id = _bootstrap_calendar(sim_date)
    tools = _get_anthropic_tools()

    system = _build_system_blocks()
    user_msg = _build_user_message(email, sim_date, scenario_id, calendar_id)

    # max_retries=3 lets the SDK absorb transient 429s on burst-y days (4+ emails
    # in close succession can momentarily exceed the 50k tok/min Haiku limit).
    # Backoff is exponential, capped well under our 300s per-turn deadline.
    client = anthropic.Anthropic(api_key=api_key, timeout=120.0, max_retries=3)

    # With continuity ON, append to the persistent chain so the model retains
    # prior emails' user/assistant/tool_result blocks. With it OFF, behave like
    # the legacy fresh-per-turn runner (for A/B comparison).
    if CONVERSATION_CONTINUITY:
        messages = _scenario_messages[scenario_id]
    else:
        messages = []
    # Snapshot length so we can roll back if this turn fails partway through.
    # Anthropic rejects malformed chains (orphan tool_use without tool_result,
    # trailing user with no assistant reply), so leaking a half-finished turn
    # would break every subsequent email in the scenario.
    _pre_turn_len = len(messages)
    messages.append({"role": "user", "content": user_msg})

    # 1-based index of this email within the scenario chain. Each preamble+body
    # we append is a string-content user message; tool_result user messages
    # carry a list. So count string-content user messages to get the email
    # position. With continuity OFF this is always 1.
    email_index = sum(
        1 for m in messages if m["role"] == "user" and isinstance(m["content"], str)
    )

    deadline = time.monotonic() + PER_TURN_TIMEOUT_S
    rounds = 0

    try:
        while True:
            rounds += 1
            if rounds > ROUND_BACKSTOP:
                raise RuntimeError(
                    f"Hit {ROUND_BACKSTOP}-round backstop on scenario {scenario_id} — "
                    "model is likely looping."
                )
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise TimeoutError(
                    f"Per-turn {PER_TURN_TIMEOUT_S:.0f}s timeout exceeded "
                    f"on scenario {scenario_id}."
                )

            # Run the API call in a worker so the deadline ACTUALLY preempts it.
            api_fut = _get_api_executor().submit(
                client.messages.create,
                model=MODEL_NAME,
                max_tokens=8192,
                system=system,
                messages=messages,
                tools=tools,
            )
            try:
                response = api_fut.result(timeout=remaining)
            except concurrent.futures.TimeoutError as exc:
                raise TimeoutError(
                    f"messages.create exceeded deadline on scenario {scenario_id}."
                ) from exc

            _record_usage(scenario_id, rounds, response, messages_in_context=len(messages))

            tool_uses = [b for b in response.content if b.type == "tool_use"]
            if not tool_uses:
                # Preserve the final assistant turn so the conversation chain
                # ends on an assistant message — required for the next email
                # to append a user message without producing two consecutive
                # user roles (which Anthropic rejects).
                messages.append({"role": "assistant", "content": response.content})
                break

            messages.append({"role": "assistant", "content": response.content})

            tool_results = []
            for tu in tool_uses:
                _stats["tool_calls"][tu.name] += 1
                raw_args = dict(tu.input)  # snapshot the model's view pre-inject
                remaining = max(1.0, deadline - time.monotonic())
                args = _inject_keys(tu.name, raw_args, scenario_id, calendar_id)
                call_failed = False
                try:
                    content, is_error = _run_async(
                        _call_tool_async(tu.name, args),
                        timeout=remaining,
                    )
                except Exception as exc:
                    _stats["tool_errors"] += 1
                    content = json.dumps({"error": str(exc)})
                    call_failed = True
                else:
                    if is_error:
                        _stats["tool_errors"] += 1
                        call_failed = True
                response_bytes, response_items = _measure_response(content)
                _log_tool_call(
                    scenario_id=scenario_id,
                    round_idx=rounds,
                    tool_name=tu.name,
                    raw_args=raw_args,
                    email_index=email_index,
                    is_error=call_failed,
                    response_bytes=response_bytes,
                    response_items=response_items,
                )
                tool_results.append({
                    "type": "tool_result",
                    "tool_use_id": tu.id,
                    "content": content,
                })

            messages.append({"role": "user", "content": tool_results})

            if response.stop_reason == "end_turn":
                break

        _stats["scenarios_run"] += 1

    except Exception as exc:
        _stats["turn_failures"] += 1
        sys.stderr.write(
            f"\n[model_runner] FAIL scenario={scenario_id} "
            f"after {rounds} round(s): {exc}\n"
        )
        tail = messages[-2:] if len(messages) >= 2 else messages
        try:
            sys.stderr.write(f"[model_runner] last-messages tail: {tail!r}\n")
        except Exception:
            pass
        # Roll the persistent chain back to its pre-turn state so the next
        # email in this scenario starts from a valid conversation prefix.
        if CONVERSATION_CONTINUITY:
            del messages[_pre_turn_len:]
        raise


# --- Scenario lifecycle hook ----------------------------------------------

def scenario_completed(scenario_id: int) -> None:
    """Drop the persistent message chain for a scenario once the engine has
    graded it. Called from engine.py after `completed_scenarios_today()`.

    Without this, the dict accumulates every scenario's full conversation for
    the entire run — eventually thousands of dead chains in memory. No-op when
    continuity is disabled."""
    _scenario_messages.pop(scenario_id, None)


# --- Shutdown + summary ----------------------------------------------------

def _shutdown() -> None:
    """Tear down MCP session, stop the loop, drain the API executor."""
    if _shutdown_signal is not None and _loop is not None and _loop.is_running():
        try:
            _loop.call_soon_threadsafe(_shutdown_signal.set)
        except RuntimeError:
            pass
        # Give the lifecycle coroutine a chance to run its finally block before
        # we stop the loop (B2 mitigation — wait for _session to clear).
        for _ in range(20):  # up to ~2s
            if _session is None:
                break
            time.sleep(0.1)

    if _loop is not None and _loop.is_running():
        try:
            _loop.call_soon_threadsafe(_loop.stop)
        except RuntimeError:
            pass
    if _loop_thread is not None and _loop_thread.is_alive():
        _loop_thread.join(timeout=5.0)
    if _api_executor is not None:
        _api_executor.shutdown(wait=False, cancel_futures=True)


def _print_summary() -> None:
    total_calls = sum(_stats["tool_calls"].values())
    scenarios = _stats["scenarios_run"]
    rounds = _stats["rounds_total"]
    in_tok = _stats["input_tokens"]
    out_tok = _stats["output_tokens"]
    total_tok = in_tok + out_tok

    print("\n=== model_runner summary ===")
    print(f"model             : {MODEL_NAME}")
    print(f"scenarios run     : {scenarios}")
    print(f"calendars created : {_stats['calendars_created']}")
    print(f"total tool calls  : {total_calls}")
    print(f"tool errors       : {_stats['tool_errors']}")
    print(f"turn failures     : {_stats['turn_failures']}")
    print(f"api rounds        : {rounds}")
    print("--- tokens ---")
    print(f"input             : {in_tok}")
    print(f"output            : {out_tok}")
    print(f"total             : {total_tok}")
    if scenarios:
        print(f"avg tokens/turn   : {total_tok / scenarios:.0f}")
        print(f"avg rounds/turn   : {rounds / scenarios:.2f}")
    if _stats["tool_calls"]:
        print("tool call breakdown:")
        for name, count in sorted(_stats["tool_calls"].items(), key=lambda x: -x[1]):
            print(f"  {name:<24} {count}")
    if TOKEN_LOG_PATH:
        print(f"per-round log     : {TOKEN_LOG_PATH}")
    if TOOL_LOG_PATH:
        print(f"per-tool log      : {TOOL_LOG_PATH}")
    print("============================\n")


# atexit is LIFO. We register shutdown FIRST so it runs LAST — that way the
# summary prints while subprocess teardown is still in flight (the user sees
# the stats before any tail-end "BrokenPipeError" noise).
atexit.register(_shutdown)
atexit.register(_print_summary)
