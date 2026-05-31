"""
harness.base — the single source of truth for driving a model over one email.

Everything that is harness-agnostic lives here so there is exactly ONE
definition of each thing and exactly ONE behavior on the live path:

  - MCP config, hidden-tools set, disallowed-tools flag string (FIX-7)
  - the canonical secretary system prompt (FIX-4)
  - build_user_message / build_prompt (the two prompt shapes harnesses need)
  - bootstrap_calendar — create/reuse the shared calendar (FIX-1)
  - parse_stream_output — claude stream-json parser with token / tool /
    compaction logging and message-id dedup (FIX-3)
  - per-scenario telemetry accumulators (tokens, compaction) for FIX-9

Concrete adapters (harness/claude_p.py, harness/codex.py) are thin consumers
of this module. model_runner.py is a backwards-compat shim over it.

Pipeline shape:
    engine.py -> harness adapter -> subprocess (claude -p) -> MCP server -> FastAPI store
"""
from __future__ import annotations

import atexit
import json
import os
import threading
from collections import defaultdict
from datetime import datetime, timezone
from typing import Any

import httpx

import bench_logger as blog
from sb.schema import Email

# --- Config (env knobs; read once at import) -------------------------------

FASTAPI_BASE = os.environ.get("FASTAPI_BASE_URL", "http://localhost:8000")
MODEL_NAME = os.environ.get("CLAUDE_MODEL", "claude-haiku-4-5")
REASONING_EFFORT = os.environ.get("CLAUDE_REASONING", "")  # "low" | "medium" | "high" | "xhigh" | "max"
PER_TURN_TIMEOUT_S = float(os.environ.get("PER_TURN_TIMEOUT_S", "300"))

TOKEN_LOG_PATH = os.environ.get("TOKEN_LOG_PATH", "token_usage.jsonl")
TOOL_LOG_PATH = os.environ.get("TOOL_LOG_PATH", "tool_calls.jsonl")

# Default ON: resume the per-scenario session across emails. CONVERSATION_CONTINUITY=0
# A/Bs against fresh-per-turn behavior. This is only the DEFAULT — the engine may
# override it per adapter via get_adapter(..., conversation_continuity=...) (FIX-8).
CONVERSATION_CONTINUITY = os.environ.get("CONVERSATION_CONTINUITY", "1") == "1"

# --- MCP config / tool surface (single definition — FIX-7) -----------------

MCP_SERVER_NAME = "secretary"

# Admin/engine-internal tools the model should NOT see.
HIDDEN_TOOLS = frozenset({
    "create_scenario", "delete_scenario", "add_scenario_email",
    "create_calendar", "delete_calendar", "health_check",
    "get_email", "list_scenarios", "get_scenario",
})

MCP_CONFIG = json.dumps({
    "mcpServers": {
        MCP_SERVER_NAME: {
            "command": "bash",
            "args": ["-c", "python -m mcp_server"],
        }
    }
})

DISALLOWED_TOOLS = ",".join(f"mcp__{MCP_SERVER_NAME}__{t}" for t in HIDDEN_TOOLS)

# --- Module-level state ----------------------------------------------------

_calendar_id: str | None = None

# 1-based email index per scenario, used for tool/usage logs. Tracked regardless
# of continuity so logs are always meaningful. Released by release_scenario().
_scenario_email_counts: dict[int, int] = defaultdict(int)

# Per-scenario telemetry for FIX-9 (compaction dimension + token totals). The
# parser accumulates here; the engine pops it at scenario completion.
_scenario_compactions: dict[int, int] = defaultdict(int)
_scenario_tokens: dict[int, dict[str, int]] = defaultdict(
    lambda: {"input_tokens": 0, "output_tokens": 0, "rounds": 0}
)
# Peak accumulated input on the previous turn of each scenario, used to detect
# Claude Code's *silent* compaction by the context drop it leaves behind.
_scenario_prev_input: dict[int, int] = defaultdict(int)
# Detection thresholds: flag compaction when this turn's accumulated context
# falls below COMPACTION_DROP_RATIO of the previous turn's peak, and that peak
# was already large (so we don't flag normal small-context variation).
_COMPACTION_DROP_RATIO = 0.6
_COMPACTION_MIN_PREV = 50_000

_token_log_lock = threading.Lock()
_tool_log_lock = threading.Lock()

_stats: dict[str, Any] = {
    "scenarios_run": 0,   # counts turns (one per run_turn call)
    "tool_calls": defaultdict(int),
    "tool_errors": 0,
    "calendars_created": 0,
    "turn_failures": 0,
    "rounds_total": 0,
    "input_tokens": 0,
    "output_tokens": 0,
}

# --- Canonical system prompt (single definition — FIX-4) -------------------

STATIC_SYSTEM_PROMPT = """You are an AI executive assistant. For each email, decide the action and call the matching tool — or call no tool. Do not narrate.

TOOLS:
- create_todo: principal must track a task with a deadline.
- create_event: a time-blocked meeting/call is proposed or confirmed.
- send_email: a direct question is asked or a response is explicitly requested.
- update_todo / update_event: an email modifies something you ALREADY created earlier in this scenario. Reschedules use update_event — never delete-then-create.
- delete_todo / delete_event: ONLY to undo a duplicate you just created by mistake. Never to "reschedule" or "cancel" — use update_* for changes; for cancellations the email just doesn't need any action.
- list_emails / list_todos / list_events / get_calendar / get_todo / get_event: ONLY when you need state you don't already remember from this conversation. Prefer NOT to call — you usually remember what you wrote.
- (no tool): FYI / newsletter / auto-confirmation / marketing / status update.

ORDER: if both an event AND a linked todo are needed, create the event first, then pass its returned id as `calendar_event_id` to create_todo.

DATETIMES: ISO 8601 with timezone (e.g. "2000-05-09T14:00:00Z"). Use the simulated date in FOR THIS EMAIL for relative phrases. Defaults: events 09:00, todo due_date 17:00, 1-hour duration.

SCENARIO_ID: copy verbatim from the LATEST FOR THIS EMAIL into every tool call.

CONTEXT: prior emails in this scenario (if any) appear earlier in this conversation along with the tool calls you made for them. Reuse that context — do NOT call list_emails for chain history you already remember.

REPLIES (send_email): recipients = incoming sender. Subject = "Re: <original>" (no double prefix). 1-3 sentences for simple Q's.

END IMMEDIATELY after the tool call(s). Do not emit confirmation text after a tool succeeds — that wastes a round.
"""

# --- Calendar bootstrap (FIX-1) --------------------------------------------

def bootstrap_calendar(sim_date: datetime) -> str:
    """Return a valid shared calendar id, creating it once per process and
    recreating it if the store was reset (e.g. uvicorn restart wiped memory).

    Module-level cache so every adapter in this process shares one calendar —
    keeps the live path harness-agnostic without threading an id through the
    HarnessAdapter interface.
    """
    global _calendar_id

    if _calendar_id is not None:
        try:
            r = httpx.get(f"{FASTAPI_BASE}/calendars/{_calendar_id}", timeout=10)
            if r.status_code == 200:
                return _calendar_id
        except Exception:
            pass
        _calendar_id = None

    aware = sim_date if sim_date.tzinfo else sim_date.replace(tzinfo=timezone.utc)
    iso = aware.isoformat().replace("+00:00", "Z")

    r = httpx.post(f"{FASTAPI_BASE}/calendars/", json={"start_date": iso}, timeout=30)
    if r.status_code != 201:
        raise RuntimeError(f"create_calendar failed: {r.status_code} {r.text}")

    data = r.json()
    if "calendar_id" not in data:
        raise RuntimeError(f"create_calendar response missing calendar_id: {data!r}")

    _calendar_id = data["calendar_id"]
    _stats["calendars_created"] += 1
    return _calendar_id


def reset_calendar_cache() -> None:
    """Drop the cached calendar id (used by tests and on hard failure)."""
    global _calendar_id
    _calendar_id = None

# --- Logging (FIX-3) -------------------------------------------------------

def _record_usage(
    scenario_id: int,
    round_idx: int,
    email_index: int,
    input_tokens: int,
    output_tokens: int,
    cache_creation_input_tokens: int = 0,
    cache_read_input_tokens: int = 0,
    stop_reason: str | None = None,
    tool_use_count: int = 0,
) -> None:
    total_input = input_tokens + cache_creation_input_tokens + cache_read_input_tokens
    _stats["input_tokens"] += total_input
    _stats["output_tokens"] += output_tokens
    _stats["rounds_total"] += 1
    # Per-scenario accumulation for FIX-9.
    st = _scenario_tokens[scenario_id]
    st["input_tokens"] += total_input
    st["output_tokens"] += output_tokens
    st["rounds"] += 1
    if cache_read_input_tokens == 0 and cache_creation_input_tokens > 0:
        _stats["cache_misses"] = _stats.get("cache_misses", 0) + 1
        blog.warn("harness", f"CACHE MISS scenario={scenario_id} email={email_index} round={round_idx} cache_write={cache_creation_input_tokens} fresh_input={input_tokens}")

    if not TOKEN_LOG_PATH:
        return
    line = json.dumps({
        "ts": datetime.now(timezone.utc).isoformat(),
        "scenario_id": scenario_id,
        "email_index": email_index,
        "round": round_idx,
        "stop_reason": stop_reason,
        "tool_uses": tool_use_count,
        "continuity": CONVERSATION_CONTINUITY,
        "input_tokens": input_tokens,
        "cache_creation_input_tokens": cache_creation_input_tokens,
        "cache_read_input_tokens": cache_read_input_tokens,
        "output_tokens": output_tokens,
    })
    try:
        with _token_log_lock, open(TOKEN_LOG_PATH, "a") as f:
            f.write(line + "\n")
    except OSError:
        pass


def _log_tool_call(
    scenario_id: int,
    round_idx: int,
    tool_name: str,
    raw_args: dict,
    email_index: int,
) -> None:
    if not TOOL_LOG_PATH:
        return
    line = json.dumps({
        "ts": datetime.now(timezone.utc).isoformat(),
        "scenario_id": scenario_id,
        "round": round_idx,
        "tool_name": tool_name,
        "args": raw_args,
        "email_index": email_index,
        "continuity": CONVERSATION_CONTINUITY,
    })
    try:
        with _tool_log_lock, open(TOOL_LOG_PATH, "a") as f:
            f.write(line + "\n")
    except OSError:
        pass

# --- Prompt shapes ---------------------------------------------------------

def build_user_message(email: Email, sim_date: datetime, scenario_id: int, calendar_id: str) -> str:
    """Just the user-turn content for one email — no system prompt.

    Used by CLI harnesses that inject the system prompt via a flag
    (claude --append-system-prompt) and by SDK harnesses that pass it as a
    separate option. Emits the concrete `calendar_id` line so the model can
    actually create events (FIX-1).
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


def build_prompt(email: Email, sim_date: datetime, scenario_id: int, calendar_id: str) -> str:
    """Full single-string prompt (system + user) for CLI harnesses that take a
    concatenated prompt argument and have no separate system-prompt flag.

    Claude does NOT use this (it injects the system prompt via
    --append-system-prompt on email 1 and carries it forward with --resume);
    provided for the harness recipe in SPRINT5_REMEDIATION §7.
    """
    return f"{STATIC_SYSTEM_PROMPT}\n\n{build_user_message(email, sim_date, scenario_id, calendar_id)}"

# --- stream-json parser (FIX-3) --------------------------------------------

def parse_stream_output(output: str, scenario_id: int, email_index: int) -> str | None:
    """Parse stream-json events from `claude -p --output-format stream-json --verbose`.

    Returns the session_id from the system/init event (for --resume). Logs
    tool_use events and per-round usage, detects compaction, and dedups
    duplicate assistant emissions by message.id — all as side effects.
    """
    session_id: str | None = None

    # Stream-json emits each assistant message multiple times (once before a
    # tool_use block is appended, once after) with byte-identical usage. Dedup
    # by message.id, keeping the LATEST emission per id so we capture the full
    # content (text + tool_use) and count usage exactly once.
    ordered_msg_ids: list[str] = []
    latest_by_id: dict[str, dict] = {}
    result_error: dict | None = None

    for raw_line in output.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        try:
            event = json.loads(line)
        except json.JSONDecodeError:
            continue

        etype = event.get("type")

        if etype == "system" and event.get("subtype") == "init":
            session_id = event.get("session_id")

        elif etype == "system" and event.get("subtype") == "compaction":
            blog.info("harness", f"COMPACTION fired — scenario={scenario_id} email_index={email_index} session={session_id}")
            _stats["compactions"] = _stats.get("compactions", 0) + 1
            _scenario_compactions[scenario_id] += 1

        elif etype == "assistant":
            msg = event.get("message", {})
            msg_id = msg.get("id") or f"_anon_{len(ordered_msg_ids)}"
            if msg_id not in latest_by_id:
                ordered_msg_ids.append(msg_id)
            latest_by_id[msg_id] = msg

        elif etype == "result":
            if not session_id:
                session_id = event.get("session_id")
            if event.get("is_error"):
                result_error = event

    turn_peak_input = 0  # max accumulated context seen this turn (for drop detection)
    for round_idx, msg_id in enumerate(ordered_msg_ids, start=1):
        msg = latest_by_id[msg_id]
        usage = msg.get("usage", {})
        content = msg.get("content", [])
        ctx = msg.get("context_management")
        if ctx:
            blog.info("harness", f"COMPACTION fired — scenario={scenario_id} email_index={email_index} context={ctx}")
            _stats["compactions"] = _stats.get("compactions", 0) + 1
            _scenario_compactions[scenario_id] += 1
        tool_uses = [b for b in content if isinstance(b, dict) and b.get("type") == "tool_use"]

        for tu in tool_uses:
            name = tu.get("name", "")
            _stats["tool_calls"][name] += 1
            _log_tool_call(
                scenario_id=scenario_id,
                round_idx=round_idx,
                tool_name=name,
                raw_args=tu.get("input", {}),
                email_index=email_index,
            )

        if usage:
            total_in = (usage.get("input_tokens", 0)
                        + usage.get("cache_creation_input_tokens", 0)
                        + usage.get("cache_read_input_tokens", 0))
            turn_peak_input = max(turn_peak_input, total_in)
            _record_usage(
                scenario_id=scenario_id,
                round_idx=round_idx,
                email_index=email_index,
                input_tokens=usage.get("input_tokens", 0),
                cache_creation_input_tokens=usage.get("cache_creation_input_tokens", 0),
                cache_read_input_tokens=usage.get("cache_read_input_tokens", 0),
                output_tokens=usage.get("output_tokens", 0),
                stop_reason=msg.get("stop_reason"),
                tool_use_count=len(tool_uses),
            )

    # Detect Claude Code's *silent* compaction (no system/compaction event in
    # print mode, verified on 2.1.158): a large drop in accumulated context
    # versus the previous turn of this scenario means the session was compacted.
    prev = _scenario_prev_input.get(scenario_id, 0)
    if (turn_peak_input > 0 and prev >= _COMPACTION_MIN_PREV
            and turn_peak_input < prev * _COMPACTION_DROP_RATIO):
        blog.info("harness", f"COMPACTION fired — scenario={scenario_id} email_index={email_index} "
                             f"context {prev:,} -> {turn_peak_input:,} tokens (detected via context drop)")
        _stats["compactions"] = _stats.get("compactions", 0) + 1
        _scenario_compactions[scenario_id] += 1
    if turn_peak_input > 0:
        _scenario_prev_input[scenario_id] = turn_peak_input

    if result_error is not None:
        raise RuntimeError(f"claude turn failed: {result_error.get('result', result_error)}")

    return session_id

# --- Per-scenario bookkeeping ----------------------------------------------

def next_email_index(scenario_id: int) -> int:
    """Increment and return the 1-based email index for a scenario."""
    _scenario_email_counts[scenario_id] += 1
    return _scenario_email_counts[scenario_id]


def scenario_telemetry(scenario_id: int) -> dict[str, int]:
    """Snapshot per-scenario telemetry (compaction count + token totals).

    Used by the engine (FIX-9) to attach a compaction/token dimension to the
    per-scenario result. Non-destructive — read before release_scenario().
    """
    tok = _scenario_tokens.get(scenario_id, {"input_tokens": 0, "output_tokens": 0, "rounds": 0})
    return {
        "compactions": _scenario_compactions.get(scenario_id, 0),
        "input_tokens": tok["input_tokens"],
        "output_tokens": tok["output_tokens"],
        "rounds": tok["rounds"],
    }


def release_scenario(scenario_id: int) -> None:
    """Drop per-scenario bookkeeping after the engine has graded the scenario."""
    _scenario_email_counts.pop(scenario_id, None)
    _scenario_compactions.pop(scenario_id, None)
    _scenario_tokens.pop(scenario_id, None)
    _scenario_prev_input.pop(scenario_id, None)


def note_turn_failure(scenario_id: int, email_index: int, exc: Exception) -> None:
    """Account + log a failed turn (called by adapters on crash/timeout)."""
    _stats["turn_failures"] += 1
    blog.error("harness", f"FAIL scenario={scenario_id} email_index={email_index}: {exc}")


def note_turn_success() -> None:
    _stats["scenarios_run"] += 1

# --- Summary on exit -------------------------------------------------------

def _print_summary() -> None:
    blog.model_summary(
        _stats,
        model_name=MODEL_NAME,
        token_log=TOKEN_LOG_PATH or "",
        tool_log=TOOL_LOG_PATH or "",
    )

atexit.register(_print_summary)
