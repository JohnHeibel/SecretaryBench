from __future__ import annotations

import asyncio
import logging
import os
import pathlib
import time
from datetime import datetime

from claude_agent_sdk import query, ClaudeAgentOptions
from claude_agent_sdk.types import AssistantMessage, ResultMessage, ToolUseBlock

from loader import Email
from harness.base import (
    TurnResult,
    _STATIC_SYSTEM_PROMPT,
    build_user_message,
    bootstrap_calendar,
    PER_TURN_TIMEOUT_S,
    ROUND_BACKSTOP,
)

logger = logging.getLogger(__name__)

_MODEL = os.environ.get("AGENT_SDK_MODEL", "claude-haiku-4-5")
_MCP_CONFIG = pathlib.Path(".claude/.mcp.json")


async def _collect(prompt: str, options: ClaudeAgentOptions) -> tuple[str | None, int | None, int | None, int | None, list[str]]:
    session_id: str | None = None
    rounds: int | None = None
    input_tokens: int | None = None
    output_tokens: int | None = None
    tool_calls: list[str] = []

    async for msg in query(prompt=prompt, options=options):
        if isinstance(msg, AssistantMessage):
            for block in msg.content:
                if isinstance(block, ToolUseBlock):
                    tool_calls.append(block.name)
        elif isinstance(msg, ResultMessage):
            session_id = msg.session_id
            rounds = msg.num_turns
            if msg.usage:
                input_tokens = msg.usage.get("input_tokens")
                output_tokens = msg.usage.get("output_tokens")

    return session_id, rounds, input_tokens, output_tokens, tool_calls


class AgentSDKHarness:
    """Harness that runs the Claude Agent SDK in-process per email turn."""

    def __init__(self) -> None:
        self._calendar_id: str | None = None
        self._sessions: dict[int, str] = {}

    def run_turn(self, email: Email, sim_date: datetime, scenario_id: int) -> TurnResult | None:
        self._calendar_id = bootstrap_calendar(sim_date, self._calendar_id)
        prompt = build_user_message(email, sim_date, scenario_id, self._calendar_id)

        options = ClaudeAgentOptions(
            model=_MODEL,
            system_prompt=_STATIC_SYSTEM_PROMPT,
            mcp_servers=_MCP_CONFIG,
            max_turns=ROUND_BACKSTOP,
            permission_mode="bypassPermissions",
            resume=self._sessions.get(scenario_id),
        )

        t0 = time.monotonic()
        try:
            session_id, rounds, input_tokens, output_tokens, tool_calls = asyncio.run(
                asyncio.wait_for(_collect(prompt, options), timeout=PER_TURN_TIMEOUT_S)
            )
        except TimeoutError:
            logger.error("scenario %d: agent-sdk timed out after %.0fs", scenario_id, PER_TURN_TIMEOUT_S)
            raise
        elapsed = time.monotonic() - t0

        if session_id:
            self._sessions[scenario_id] = session_id

        return TurnResult(
            scenario_id=scenario_id,
            elapsed_s=elapsed,
            rounds=rounds,
            input_tokens=input_tokens,
            output_tokens=output_tokens,
            tool_calls=tool_calls,
        )

    def scenario_completed(self, scenario_id: int) -> None:
        self._sessions.pop(scenario_id, None)

    def shutdown(self) -> None:
        pass
