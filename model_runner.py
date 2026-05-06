"""
model_runner.py — AI lane for SecretaryBench.

Drives an LLM through the MCP server (stdio subprocess) once per email.
The model has access to every tool the MCP server exposes (22 today).

Architecture:

    engine.py (sync)
        │  run_model_turn(email, sim_date, scenario_id)
        ▼
    model_runner.py
        │  - Anthropic SDK (model layer)
        │  - MCP stdio client (tool layer, harness-agnostic)
        ▼
    mcp_server (subprocess, stdio)
        │  HTTP
        ▼
    FastAPI store

Key decisions (see modelRunner.md for the why):

- One MCP session per benchmark run. Lazy-init on first call to
  run_model_turn, torn down at process exit. We do NOT spawn the
  subprocess per email — that would cost ~109 spawns per run.
- A dedicated asyncio event loop runs in a daemon thread; the sync
  engine submits work via run_coroutine_threadsafe and blocks for
  the result. This lets the MCP session live across many sync calls.
- One shared calendar bootstrapped via the MCP `create_calendar`
  tool on session open. Its id is injected into the system prompt
  so the model uses it. We never call FastAPI directly from here —
  every store mutation goes through MCP.
- No fixed cap on tool rounds (one email may legitimately need a
  multi-step chain, e.g. create_event then create_todo linking it).
  A 25-round backstop and a 300s per-turn timeout are the only safety
  bounds; both crash the run loudly with the scenario_id logged.
- MODEL is hardcoded to claude-sonnet-4-6 for now. Swapping to a
  different provider later means dropping in litellm; the rest of
  this file (MCP layer, session lifecycle, stats) doesn't change.

Run summary: an atexit hook prints per-tool call counts, scenarios
processed, calendars created, and errors when the process exits.
"""
from __future__ import annotations

import asyncio
import atexit
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
_shutdown_signal: asyncio.Event | None = None # set on atexit to tear down

_calendar_id: str | None = None
_cached_anthropic_tools: list[dict] | None = None

_stats: dict[str, Any] = {
    "scenarios_run": 0,
    "tool_calls": defaultdict(int),
    "tool_errors": 0,
    "calendars_created": 0,
    "turn_failures": 0,
}

MODEL_NAME = "claude-sonnet-4-6"
PER_TURN_TIMEOUT_S = 300.0
ROUND_BACKSTOP = 25
SESSION_OPEN_TIMEOUT_S = 30.0


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
    """Submit a coroutine to the daemon loop and block for its result."""
    loop = _ensure_loop()
    fut = asyncio.run_coroutine_threadsafe(coro, loop)
    return fut.result(timeout=timeout)


# --- MCP session lifecycle -------------------------------------------------

async def _session_lifecycle() -> None:
    """Long-running coroutine that owns the MCP subprocess + session.

    Holds the async context managers open until _shutdown_signal is set,
    so the session survives across all sync run_model_turn calls.
    """
    global _session, _shutdown_signal
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


def _ensure_session() -> None:
    """Idempotent: opens the MCP session on the daemon loop if not open."""
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
    ]


async def _call_tool_async(name: str, args: dict) -> str:
    """Call an MCP tool and return its content as a single string."""
    result = await _session.call_tool(name, args)
    parts: list[str] = []
    for block in result.content:
        text = getattr(block, "text", None)
        if text is not None:
            parts.append(text)
        else:
            parts.append(str(block))
    payload = "\n".join(parts) if parts else "{}"
    if getattr(result, "isError", False):
        return json.dumps({"error": payload})
    return payload


def _get_anthropic_tools() -> list[dict]:
    """Lazy-cache the tool list — schemas don't change mid-run."""
    global _cached_anthropic_tools
    if _cached_anthropic_tools is None:
        _cached_anthropic_tools = _run_async(_list_tools_async(), timeout=30)
    return _cached_anthropic_tools


# --- Calendar bootstrap ----------------------------------------------------

def _bootstrap_calendar(sim_date: datetime) -> str:
    """Create one shared calendar via the MCP create_calendar tool."""
    global _calendar_id
    if _calendar_id is not None:
        return _calendar_id

    aware = sim_date if sim_date.tzinfo else sim_date.replace(tzinfo=timezone.utc)
    iso = aware.isoformat().replace("+00:00", "Z")

    raw = _run_async(_call_tool_async("create_calendar", {"start_date": iso}), timeout=30)
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

You will receive ONE email at a time. Decide on the right action(s) using the tools available to you.

POSSIBLE ACTIONS PER EMAIL:
1. CREATE A TODO    — the email asks you to track a task, follow up, or complete an action item.
2. SCHEDULE AN EVENT — the email involves setting up a meeting or time-based event.
3. SEND A REPLY     — the email explicitly requires a written response back to the sender.
4. DO NOTHING       — the email is purely informational (FYI, newsletter, automated notification,
                      confirmation that needs no reply or action). Call no tools.

You can call MULTIPLE tools in a single email if it genuinely needs them — for example:
create the calendar event first, then create a todo that links to it via calendar_event_id.
Order matters: an event must exist before a todo can reference it.

REQUIRED FIELDS:
- scenario_id   = {scenario_id}                (use this exact value for every todo and event)
- calendar_id   = "{calendar_id}"               (use this exact value when creating events)
- All datetimes are ISO 8601 with timezone, e.g. "2000-01-15T10:00:00Z".
- Today's simulated date is {sim_date_str} (ISO: {sim_date_iso}).

GUIDELINES:
- If the email is informational, do nothing. Over-acting costs points.
- Do not create duplicates of todos/events you've already created in this turn.
- Default event duration is 1 hour unless the email specifies otherwise.
- You also have read tools (list_*, get_*) and an api-reference resource if you need them,
  but most emails should be solvable from the email text alone.
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

    Side-effects (todos, events, replies) land in the FastAPI store via
    MCP tool calls. Returns nothing — the grader reads results back.
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

    client = anthropic.Anthropic(api_key=api_key, timeout=120.0)
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
            if time.monotonic() > deadline:
                raise TimeoutError(
                    f"Per-turn {PER_TURN_TIMEOUT_S:.0f}s timeout exceeded "
                    f"on scenario {scenario_id}."
                )

            response = client.messages.create(
                model=MODEL_NAME,
                max_tokens=8192,
                system=system,
                messages=messages,
                tools=tools,
            )

            tool_uses = [b for b in response.content if b.type == "tool_use"]
            if not tool_uses:
                break

            messages.append({"role": "assistant", "content": response.content})

            tool_results = []
            for tu in tool_uses:
                _stats["tool_calls"][tu.name] += 1
                remaining = max(1.0, deadline - time.monotonic())
                try:
                    content = _run_async(
                        _call_tool_async(tu.name, dict(tu.input)),
                        timeout=remaining,
                    )
                except Exception as exc:
                    _stats["tool_errors"] += 1
                    content = json.dumps({"error": str(exc)})
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
    """Tear down the MCP session and stop the daemon loop. Safe to call twice."""
    if _shutdown_signal is not None and _loop is not None and _loop.is_running():
        try:
            _loop.call_soon_threadsafe(_shutdown_signal.set)
        except RuntimeError:
            pass
    if _loop is not None and _loop.is_running():
        try:
            _loop.call_soon_threadsafe(_loop.stop)
        except RuntimeError:
            pass
    if _loop_thread is not None and _loop_thread.is_alive():
        _loop_thread.join(timeout=5.0)


def _print_summary() -> None:
    total_calls = sum(_stats["tool_calls"].values())
    print("\n=== model_runner summary ===")
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


# Summary registers first → runs LAST (atexit is LIFO). We want shutdown to
# happen first so the summary prints after the subprocess is reaped cleanly.
atexit.register(_print_summary)
atexit.register(_shutdown)
