from __future__ import annotations

"""
harness.py — Harness Abstraction Layer

Defines the HarnessAdapter interface and concrete implementations for each
supported harness (Claude Code, Codex). The engine talks to this layer —
swapping harness or model is one CLI/env value, zero changes to engine,
grader, or MCP.

How it fits in the pipeline:
    engine.py -> harness.py -> subprocess (claude / codex) -> MCP server -> FastAPI

Adding a new harness:
    1. Subclass HarnessAdapter
    2. Implement the 4 methods
    3. Register it in HARNESS_REGISTRY at the bottom of this file

TESTING NOTES:
    The logic in this file (adapter instantiation, session management, message
    formatting, session_id parsing) can be smoke-tested without any external
    dependencies by running:

        python3 -c "
        from harness import get_adapter, _build_user_message, _extract_session_id
        from loader import Email
        from datetime import datetime

        adapter = get_adapter('claude-code', model='claude-haiku-4-5')
        adapter.start_session(42)

        email = Email(1, 'Meeting request', 'Can we meet Thursday?', 'boss@corp.com', ['me@corp.com'])
        msg = _build_user_message(email, datetime(2025, 5, 1), 42)
        print(msg)

        sid = _extract_session_id('{\"type\":\"system\",\"subtype\":\"init\",\"session_id\":\"abc-123\"}')
        print('session_id:', sid)

        adapter.end_session(42)
        "

    Full end-to-end testing (run_turn) was not possible locally due to two limitations:
        1. The `claude` CLI was not installed on the development machine.
        2. The FastAPI server must be running for the MCP server to connect to.
    Full testing requires the whole bench to be running — FastAPI up via
    `uvicorn app.main:app --reload` and `claude` CLI installed.
"""

import json
import os
import subprocess
from abc import ABC, abstractmethod
from datetime import datetime
from typing import Optional

from loader import Email


_MCP_SERVER_NAME = "secretary"
_MCP_CONFIG = json.dumps({
    "mcpServers": {
        _MCP_SERVER_NAME: {
            "command": "bash",
            "args": ["-c", "python -m mcp_server"],
        }
    }
})

_HIDDEN_TOOLS = frozenset({
    "create_scenario", "delete_scenario", "add_scenario_email",
    "create_calendar", "delete_calendar", "health_check",
    "get_email", "list_scenarios", "get_scenario",
})
_DISALLOWED_TOOLS = ",".join(f"mcp__{_MCP_SERVER_NAME}__{t}" for t in _HIDDEN_TOOLS)


class HarnessAdapter(ABC):
    """Interface every harness adapter must implement.

    The engine calls these 4 methods in order for each scenario:
      1. start_session()   — scenario activates
      2. run_turn()        — one email served
      3. resume_session()  — emails 2..N re-attach context
      4. end_session()     — scenario completes
    """

    @abstractmethod
    def start_session(self, scenario_id: int) -> None: ...

    @abstractmethod
    def run_turn(self, email: Email, sim_date: datetime, scenario_id: int) -> None: ...

    @abstractmethod
    def resume_session(self, scenario_id: int) -> None: ...

    @abstractmethod
    def end_session(self, scenario_id: int) -> None: ...


class ClaudeCodeAdapter(HarnessAdapter):
    """Drives Claude Code non-interactively as the harness.

    Each scenario maps to one Claude Code session. Emails 2..N resume the same
    session so conversation accumulates and compaction can fire automatically.
    Supports OpenRouter via ANTHROPIC_BASE_URL + ANTHROPIC_API_KEY env vars.
    """

    def __init__(
        self,
        model: str = "claude-haiku-4-5",
        reasoning: Optional[str] = None,          # "low" / "medium" / "high" -> --effort
        api_base: Optional[str] = None,            # OpenRouter or custom base URL
        openrouter_key: Optional[str] = None,
        system_prompt: Optional[str] = None,
        per_turn_timeout: float = 300.0,
        conversation_continuity: bool = True,      # False = fresh session per email
    ):
        self.model = model
        self.reasoning = reasoning
        self.api_base = api_base
        self.openrouter_key = openrouter_key
        self.system_prompt = system_prompt or _DEFAULT_SYSTEM_PROMPT
        self.per_turn_timeout = per_turn_timeout
        self.conversation_continuity = conversation_continuity
        self._sessions: dict[int, Optional[str]] = {}

    def start_session(self, scenario_id: int) -> None:
        self._sessions[scenario_id] = None

    def run_turn(self, email: Email, sim_date: datetime, scenario_id: int) -> None:
        existing_session = self._sessions.get(scenario_id) if self.conversation_continuity else None

        cmd = [
            "claude", "-p",
            "--output-format", "stream-json",
            "--verbose",
            "--tools", "",
            "--permission-mode", "bypassPermissions",
            "--mcp-config", _MCP_CONFIG,
            "--strict-mcp-config",
            "--disallowed-tools", _DISALLOWED_TOOLS,
        ]

        if existing_session is None:
            cmd += ["--model", self.model, "--append-system-prompt", self.system_prompt]
        else:
            cmd += ["--resume", existing_session]

        if self.reasoning:
            cmd += ["--effort", self.reasoning]

        cmd.append(_build_user_message(email, sim_date, scenario_id))

        env = os.environ.copy()
        if self.api_base:
            env["ANTHROPIC_BASE_URL"] = self.api_base
        if self.openrouter_key:
            env["ANTHROPIC_API_KEY"] = self.openrouter_key

        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, env=env)
        try:
            stdout, stderr = proc.communicate(timeout=self.per_turn_timeout)
        except subprocess.TimeoutExpired:
            proc.kill()
            proc.communicate()
            raise TimeoutError(f"Harness timeout ({self.per_turn_timeout}s) on scenario {scenario_id}")

        if proc.returncode != 0:
            raise RuntimeError(f"claude exited {proc.returncode} on scenario {scenario_id}: {stderr.strip()}")

        session_id = _extract_session_id(stdout)
        if self.conversation_continuity and session_id:
            self._sessions[scenario_id] = session_id

    def resume_session(self, scenario_id: int) -> None:
        pass  # handled inside run_turn via --resume

    def end_session(self, scenario_id: int) -> None:
        self._sessions.pop(scenario_id, None)


class CodexAdapter(HarnessAdapter):
    """Stub for OpenAI Codex harness. MCP config is identical to ClaudeCodeAdapter."""

    def __init__(self, model: str = "o4-mini", api_key: Optional[str] = None):
        self.model = model
        self.api_key = api_key
        self._sessions: dict[int, Optional[str]] = {}

    def start_session(self, scenario_id: int) -> None:
        self._sessions[scenario_id] = None

    def run_turn(self, email: Email, sim_date: datetime, scenario_id: int) -> None:
        raise NotImplementedError("Codex adapter not yet implemented. Use --harness claude-code.")

    def resume_session(self, scenario_id: int) -> None:
        raise NotImplementedError("Codex adapter not yet implemented.")

    def end_session(self, scenario_id: int) -> None:
        self._sessions.pop(scenario_id, None)


def _build_user_message(email: Email, sim_date: datetime, scenario_id: int) -> str:
    recipients = ", ".join(email.recipients) if email.recipients else "(none)"
    sim_date_str = sim_date.strftime("%B %d, %Y")
    sim_date_iso = sim_date.strftime("%Y-%m-%dT%H:%M:%SZ")
    return (
        f"FOR THIS EMAIL:\n"
        f"- scenario_id = {scenario_id}\n"
        f"- Today's simulated date is {sim_date_str} (ISO: {sim_date_iso}).\n\n"
        f"From: {email.sender}\n"
        f"To: {recipients}\n"
        f"Subject: {email.subject}\n"
        f"Date: {sim_date.strftime('%Y-%m-%d')}\n\n"
        f"{email.body}"
    )


def _extract_session_id(output: str) -> Optional[str]:
    """Parse session_id from claude's stream-json output."""
    session_id: str | None = None
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
        elif etype == "result" and not session_id:
            session_id = event.get("session_id")
    return session_id


_DEFAULT_SYSTEM_PROMPT = """You are an AI executive assistant. For each email, decide the action and call the matching tool — or call no tool. Do not narrate.

TOOLS:
- create_todo: principal must track a task with a deadline.
- create_event: a time-blocked meeting/call is proposed or confirmed.
- send_email: a direct question is asked or a response is explicitly requested.
- update_todo / update_event: an email modifies something you ALREADY created earlier in this scenario.
- (no tool): FYI / newsletter / auto-confirmation / marketing / status update.

When in doubt, do nothing. Over-acting is the most common failure.

SCENARIO_ID: copy verbatim from FOR THIS EMAIL into every tool call.
DATETIMES: ISO 8601 with timezone (e.g. "2000-05-09T14:00:00Z").
"""


HARNESS_REGISTRY: dict[str, type[HarnessAdapter]] = {
    "claude-code": ClaudeCodeAdapter,
    "codex": CodexAdapter,
}


def get_adapter(harness: str, **kwargs) -> HarnessAdapter:
    """Instantiate the right adapter by name. kwargs flow directly to __init__."""
    cls = HARNESS_REGISTRY.get(harness)
    if cls is None:
        raise ValueError(f"Unknown harness {harness!r}. Available: {list(HARNESS_REGISTRY.keys())}")
    return cls(**kwargs)
