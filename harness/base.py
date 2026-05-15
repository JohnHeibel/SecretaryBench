from __future__ import annotations

import os
from dataclasses import dataclass, field
from datetime import datetime
from typing import Protocol, runtime_checkable

from loader import Email


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

PER_TURN_TIMEOUT_S = float(os.environ.get("PER_TURN_TIMEOUT_S", "300"))
ROUND_BACKSTOP = int(os.environ.get("ROUND_BACKSTOP", "25"))


def build_user_message(email: Email, sim_date: datetime, scenario_id: int, calendar_id: str) -> str:
    """Build just the user-turn content for one email — no system prompt.

    Use this for SDK harnesses that pass the system prompt separately via
    options (e.g. AgentSDKHarness passes system_prompt=_STATIC_SYSTEM_PROMPT
    to the SDK, then sends this as the user message).
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
    """Build the full prompt for CLI harnesses (system prompt + user message).

    Combines _STATIC_SYSTEM_PROMPT + per-email preamble + email body into a
    single string for passing as the -p argument to claude or codex.
    """
    return f"{_STATIC_SYSTEM_PROMPT}\n\n{build_user_message(email, sim_date, scenario_id, calendar_id)}"


STORE_BASE_URL = os.environ.get("STORE_BASE_URL", "http://127.0.0.1:8000")


def bootstrap_calendar(sim_date: datetime, cached_id: str | None = None) -> str:
    """Get or create the shared benchmark calendar. Call on first run_turn.

    If cached_id is provided, verifies it still exists in the store (guards
    against uvicorn restarts wiping the in-memory store). Creates a fresh
    calendar if the cached one is gone or if cached_id is None.
    """
    import urllib.request, urllib.error, json as _json

    if cached_id is not None:
        try:
            with urllib.request.urlopen(f"{STORE_BASE_URL}/calendars/{cached_id}", timeout=10) as r:
                if r.status == 200:
                    return cached_id
        except urllib.error.HTTPError:
            pass  # 404 — store was reset, fall through to recreate

    aware = sim_date if sim_date.tzinfo else sim_date.replace(tzinfo=timezone.utc)
    body = _json.dumps({"start_date": aware.isoformat().replace("+00:00", "Z")}).encode()
    req = urllib.request.Request(f"{STORE_BASE_URL}/calendars/", data=body, headers={"Content-Type": "application/json"}, method="POST")
    with urllib.request.urlopen(req, timeout=30) as r:
        data = _json.loads(r.read())
    return data["calendar_id"]


@dataclass
class TurnResult:
    scenario_id: int
    elapsed_s: float
    rounds: int | None = None
    input_tokens: int | None = None
    output_tokens: int | None = None
    tool_calls: list[str] = field(default_factory=list)


@runtime_checkable
class ModelHarness(Protocol):
    def run_turn(self, email: Email, sim_date: datetime, scenario_id: int) -> TurnResult | None: ...
    def scenario_completed(self, scenario_id: int) -> None: ...
    def shutdown(self) -> None: ...
