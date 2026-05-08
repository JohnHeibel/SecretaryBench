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

_stats: dict[str, Any] = {
    "scenarios_run": 0,
    "tool_calls": defaultdict(int),
    "tool_errors": 0,
    "calendars_created": 0,
    "turn_failures": 0,
}

MODEL_NAME = "claude-haiku-4-5"
PER_TURN_TIMEOUT_S = 300.0
ROUND_BACKSTOP = 25
SESSION_OPEN_TIMEOUT_S = 30.0

# Tools the model should NOT see. The MCP server still exposes them (other
# harnesses keep working); we just don't pass them in the `tools=` list to
# the model. Admin/destructive tools cause benchmark corruption (the model
# fabricates scenarios when it sees a 404) or waste tokens.
HIDDEN_TOOLS = frozenset({
    "create_scenario",
    "delete_scenario",
    "add_scenario_email",
    "create_calendar",
    "delete_calendar",
    "delete_todo",
    "delete_event",
    "health_check",
})

# Tools where the runner FORCES the foreign-key field to the correct value
# before forwarding to MCP. Small models (esp. Haiku) garble the 10-digit
# scenario_id hash and the calendar UUID; the benchmark tests "can the model
# decide what action to take", not "can it copy a number". We let the model
# pass whatever it wants — we overwrite before submission.
_FORCE_SCENARIO_ID = frozenset({
    "create_todo", "create_event", "send_email", "update_event",
})
_FORCE_CALENDAR_ID = frozenset({
    "create_event", "update_event", "list_events", "get_event", "get_calendar",
})


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


def _get_anthropic_tools() -> list[dict]:
    """Lazy-cache the (filtered) tool list."""
    global _cached_anthropic_tools
    if _cached_anthropic_tools is None:
        _cached_anthropic_tools = _run_async(_list_tools_async(), timeout=30)
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

def _build_system_prompt(scenario_id: int, calendar_id: str, sim_date: datetime) -> str:
    sim_date_str = sim_date.strftime("%B %d, %Y")
    sim_date_iso = sim_date.strftime("%Y-%m-%dT%H:%M:%SZ")
    return f"""You are an AI executive assistant managing a professional's schedule, todos, and email.

You will receive ONE email at a time. Decide on the right action(s) using the tools available.

POSSIBLE ACTIONS PER EMAIL:
1. CREATE A TODO     — the email asks you to track a task or follow up.
2. SCHEDULE AN EVENT — the email involves setting up a meeting or time-based event.
3. SEND A REPLY      — the email explicitly requires a written response.
4. DO NOTHING        — the email is purely informational (FYI, newsletter, auto-notification,
                       confirmation needing no reply or action). Call no tools.

Multi-tool actions are allowed when an email genuinely needs them. ORDER MATTERS:
create the calendar event FIRST, then create a todo that references it via
`calendar_event_id` (the field on `create_todo`). The store rejects todos that
reference a non-existent event.

═══════════════════════════════════════════════════════════
CRITICAL — SCENARIO ID:
    scenario_id = {scenario_id}

Use this EXACT integer when creating todos, events, or sent emails.
Do NOT substitute, modify, round, or invent your own scenario_id. Copy this
number verbatim into every tool call. If you receive a 404 error referencing
a scenario_id, the fix is to use the value ABOVE — never call any scenario-
management tool to "create" or "fix" the scenario; those tools are not
available to you and the scenario is already set up correctly.
═══════════════════════════════════════════════════════════

REQUIRED FIELDS:
- scenario_id  = {scenario_id}                  (the value above — exact)
- calendar_id  = "{calendar_id}"                 (use only when creating events)
- All datetimes are ISO 8601 with timezone, e.g. "2000-01-15T10:00:00Z".
- Today's simulated date is {sim_date_str} (ISO: {sim_date_iso}).

CHAIN CONTEXT:
This email may be a follow-up in a multi-email chain. Earlier emails for the
same scenario_id are already in the store. If the email reads like a reply
or follow-up and you need prior context, you may call `list_emails` (then
filter for entries whose scenario_id matches the value above) before acting.

GUIDELINES:
- If the email is informational, do nothing. Over-acting costs points.
- Do not duplicate work you've already done in this turn.
- Default event duration is 1 hour unless the email specifies otherwise.
"""


def _build_user_message(email: Email, sim_date: datetime) -> str:
    sim_date_str = sim_date.strftime("%B %d, %Y")
    recipients = ", ".join(email.recipients) if email.recipients else "(no recipients)"
    return (
        f"Email received on {sim_date_str}:\n\n"
        f"From: {email.sender}\n"
        f"To: {recipients}\n"
        f"Subject: {email.subject}\n\n"
        f"{email.body}\n\n"
        "Handle this email appropriately."
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

    system = _build_system_prompt(scenario_id, calendar_id, sim_date)
    user_msg = _build_user_message(email, sim_date)

    # max_retries=0 prevents the SDK from sleeping through 429s past our deadline.
    client = anthropic.Anthropic(api_key=api_key, timeout=120.0, max_retries=0)
    messages: list[dict] = [{"role": "user", "content": user_msg}]
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

            tool_uses = [b for b in response.content if b.type == "tool_use"]
            if not tool_uses:
                break

            messages.append({"role": "assistant", "content": response.content})

            tool_results = []
            for tu in tool_uses:
                _stats["tool_calls"][tu.name] += 1
                remaining = max(1.0, deadline - time.monotonic())
                args = _inject_keys(tu.name, dict(tu.input), scenario_id, calendar_id)
                try:
                    content, is_error = _run_async(
                        _call_tool_async(tu.name, args),
                        timeout=remaining,
                    )
                except Exception as exc:
                    _stats["tool_errors"] += 1
                    content = json.dumps({"error": str(exc)})
                else:
                    if is_error:
                        _stats["tool_errors"] += 1
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
        raise


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
    print("\n=== model_runner summary ===")
    print(f"model             : {MODEL_NAME}")
    print(f"scenarios run     : {_stats['scenarios_run']}")
    print(f"calendars created : {_stats['calendars_created']}")
    print(f"total tool calls  : {total_calls}")
    print(f"tool errors       : {_stats['tool_errors']}")
    print(f"turn failures     : {_stats['turn_failures']}")
    if _stats["tool_calls"]:
        print("tool call breakdown:")
        for name, count in sorted(_stats["tool_calls"].items(), key=lambda x: -x[1]):
            print(f"  {name:<24} {count}")
    print("============================\n")


# atexit is LIFO. We register shutdown FIRST so it runs LAST — that way the
# summary prints while subprocess teardown is still in flight (the user sees
# the stats before any tail-end "BrokenPipeError" noise).
atexit.register(_shutdown)
atexit.register(_print_summary)
