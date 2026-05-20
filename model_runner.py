"""
model_runner.py — AI lane for SecretaryBench.

Drives claude (Claude Code CLI) non-interactively once per email via
subprocess. Per-scenario claude session IDs accumulate conversation across
emails; claude's built-in compaction handles context-window overflow.

Architecture:

    engine.py (sync)
        │  run_model_turn(email, sim_date, scenario_id)
        ▼
    model_runner.py
        │  - claude -p subprocess (model + tool-call loop; MCP managed by claude)
        │  - httpx (calendar bootstrap directly against FastAPI)
        ▼
    FastAPI store  ←  mcp_server subprocess (spawned by claude)
"""
from __future__ import annotations

import atexit
import json
import os
import subprocess
import threading
from collections import defaultdict
from datetime import datetime, timezone
from typing import Any

import httpx

import bench_logger as blog
from loader import Email

# --- Config ---------------------------------------------------------------

FASTAPI_BASE = os.environ.get("FASTAPI_BASE_URL", "http://localhost:8000")
MODEL_NAME = os.environ.get("CLAUDE_MODEL", "claude-haiku-4-5")
PER_TURN_TIMEOUT_S = 300.0

TOKEN_LOG_PATH = os.environ.get("TOKEN_LOG_PATH", "token_usage.jsonl")
TOOL_LOG_PATH = os.environ.get("TOOL_LOG_PATH", "tool_calls.jsonl")

# Default ON: resume the per-scenario claude session across emails.
# Set CONVERSATION_CONTINUITY=0 to A/B against fresh-per-turn behavior.
CONVERSATION_CONTINUITY = os.environ.get("CONVERSATION_CONTINUITY", "1") == "1"

_MCP_SERVER_NAME = "secretary"

# Admin/engine-internal tools the model should NOT see.
HIDDEN_TOOLS = frozenset({
    "create_scenario", "delete_scenario", "add_scenario_email",
    "create_calendar", "delete_calendar", "health_check",
    "get_email", "list_scenarios", "get_scenario",
})

_MCP_CONFIG = json.dumps({
    "mcpServers": {
        _MCP_SERVER_NAME: {
            "command": "bash",
            "args": ["-c", "uv run python -m mcp_server"],
        }
    }
})

_DISALLOWED_TOOLS = ",".join(f"mcp__{_MCP_SERVER_NAME}__{t}" for t in HIDDEN_TOOLS)

# --- Module-level state ---------------------------------------------------

_calendar_id: str | None = None

# Per-scenario claude session IDs. None = no session yet (email 1 starts fresh).
# Cleared by scenario_completed() after grading.
_scenario_sessions: dict[int, str | None] = defaultdict(lambda: None)

# 1-based email index per scenario for _log_tool_call, tracked regardless of
# CONVERSATION_CONTINUITY so tool logs are always meaningful.
_scenario_email_counts: dict[int, int] = defaultdict(int)

_token_log_lock = threading.Lock()
_tool_log_lock = threading.Lock()

_stats: dict[str, Any] = {
    "scenarios_run": 0,   # counts turns (one per run_model_turn call)
    "tool_calls": defaultdict(int),
    "tool_errors": 0,
    "calendars_created": 0,
    "turn_failures": 0,
    "rounds_total": 0,
    "input_tokens": 0,
    "output_tokens": 0,
}

# --- System prompt --------------------------------------------------------

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

# --- Calendar bootstrap ---------------------------------------------------

def _bootstrap_calendar(sim_date: datetime) -> str:
    """Return a valid shared calendar id, recreating if the store was reset."""
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

# --- Logging --------------------------------------------------------------

def _record_usage(
    scenario_id: int,
    round_idx: int,
    input_tokens: int,
    output_tokens: int,
    stop_reason: str | None = None,
    tool_use_count: int = 0,
) -> None:
    _stats["input_tokens"] += input_tokens
    _stats["output_tokens"] += output_tokens
    _stats["rounds_total"] += 1

    if not TOKEN_LOG_PATH:
        return
    line = json.dumps({
        "ts": datetime.now(timezone.utc).isoformat(),
        "scenario_id": scenario_id,
        "round": round_idx,
        "stop_reason": stop_reason,
        "tool_uses": tool_use_count,
        "continuity": CONVERSATION_CONTINUITY,
        "input_tokens": input_tokens,
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

# --- User message ---------------------------------------------------------

def _build_user_message(email: Email, sim_date: datetime, scenario_id: int, calendar_id: str) -> str:
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

# --- stream-json parser ---------------------------------------------------

def _parse_stream_output(output: str, scenario_id: int, email_index: int) -> str | None:
    """Parse stream-json events from claude -p --verbose.

    Returns the session_id from the system/init event for use with --resume.
    Logs tool_use events and per-round usage as side-effects.
    """
    session_id: str | None = None
    round_idx = 0

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
            blog.info("model_runner", f"COMPACTION fired — scenario={scenario_id} email_index={email_index} session={session_id}")
            _stats["compactions"] = _stats.get("compactions", 0) + 1

        elif etype == "assistant":
            round_idx += 1
            msg = event.get("message", {})
            usage = msg.get("usage", {})
            content = msg.get("content", [])
            ctx = msg.get("context_management")
            if ctx:
                blog.info("model_runner", f"COMPACTION fired — scenario={scenario_id} email_index={email_index} context={ctx}")
                _stats["compactions"] = _stats.get("compactions", 0) + 1
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
                _record_usage(
                    scenario_id=scenario_id,
                    round_idx=round_idx,
                    input_tokens=usage.get("input_tokens", 0) + usage.get("cache_creation_input_tokens", 0) + usage.get("cache_read_input_tokens", 0),
                    output_tokens=usage.get("output_tokens", 0),
                    stop_reason=msg.get("stop_reason"),
                    tool_use_count=len(tool_uses),
                )

        elif etype == "result":
            if not session_id:
                session_id = event.get("session_id")
            if event.get("is_error"):
                raise RuntimeError(f"claude turn failed: {event.get('result', event)}")

    return session_id

# --- Public interface -----------------------------------------------------

def run_model_turn(email: Email, sim_date: datetime, scenario_id: int = 0) -> None:
    """Send one email to claude and let it act through MCP tools.

    Blocks until the claude subprocess finishes. Side-effects (todos, events,
    replies) land in the FastAPI store. Returns nothing — the grader reads
    results back from the store.
    """
    calendar_id = _bootstrap_calendar(sim_date)
    user_msg = _build_user_message(email, sim_date, scenario_id, calendar_id)

    _scenario_email_counts[scenario_id] += 1
    email_index = _scenario_email_counts[scenario_id]
    existing_session = _scenario_sessions[scenario_id] if CONVERSATION_CONTINUITY else None

    cmd = [
        "claude", "-p",
        "--output-format", "stream-json",
        "--verbose",
        "--tools", "",
        "--permission-mode", "bypassPermissions",
        "--mcp-config", _MCP_CONFIG,
        "--disallowed-tools", _DISALLOWED_TOOLS,
    ]

    if existing_session is None:
        cmd += ["--model", MODEL_NAME, "--append-system-prompt", _STATIC_SYSTEM_PROMPT]
    else:
        cmd += ["--resume", existing_session]

    effort = os.environ.get("CLAUDE_REASONING")
    if effort:
        cmd += ["--effort", effort]

    cmd.append(user_msg)

    try:
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        try:
            stdout, stderr = proc.communicate(timeout=PER_TURN_TIMEOUT_S)
        except subprocess.TimeoutExpired:
            proc.kill()
            stdout, stderr = proc.communicate()
            raise TimeoutError(f"Per-turn {PER_TURN_TIMEOUT_S:.0f}s timeout on scenario {scenario_id}.")

        if proc.returncode != 0:
            raise RuntimeError(f"claude exited {proc.returncode} on scenario {scenario_id}: {stderr.strip()}")

        session_id = _parse_stream_output(stdout, scenario_id, email_index)

        if CONVERSATION_CONTINUITY and session_id:
            _scenario_sessions[scenario_id] = session_id

        _stats["scenarios_run"] += 1

    except Exception as exc:
        _stats["turn_failures"] += 1
        blog.error("model_runner", f"FAIL scenario={scenario_id} email_index={email_index}: {exc}")
        raise


def scenario_completed(scenario_id: int) -> None:
    """Release the claude session after the engine has graded the scenario.

    No-op when continuity is disabled.
    """
    _scenario_sessions.pop(scenario_id, None)
    _scenario_email_counts.pop(scenario_id, None)

# --- Summary on exit ------------------------------------------------------

def _print_summary() -> None:
    blog.model_summary(
        _stats,
        model_name=MODEL_NAME,
        token_log=TOKEN_LOG_PATH or "",
        tool_log=TOOL_LOG_PATH or "",
    )

atexit.register(_print_summary)
