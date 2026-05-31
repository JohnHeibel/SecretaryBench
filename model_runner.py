"""
model_runner.py — backwards-compat shim.

The "drive claude for one email" logic now lives in the harness/ package
(harness.base + harness.claude_p) so there is exactly ONE behavior on the live
path. This module survives only so existing imports keep working:

    from model_runner import run_model_turn, scenario_completed

Both delegate to a process-wide ClaudeCodeAdapter built from the same env knobs
the package reads (CLAUDE_MODEL, CLAUDE_REASONING, CONVERSATION_CONTINUITY, ...).
The engine prefers the harness adapter directly; this path is the fallback when
no adapter was constructed.
"""
from __future__ import annotations

from datetime import datetime

from loader import Email
from harness import base
from harness.claude_p import ClaudeCodeAdapter

# Re-export so any straggling `from model_runner import X` keeps resolving.
MODEL_NAME = base.MODEL_NAME
REASONING_EFFORT = base.REASONING_EFFORT
CONVERSATION_CONTINUITY = base.CONVERSATION_CONTINUITY
TOKEN_LOG_PATH = base.TOKEN_LOG_PATH
TOOL_LOG_PATH = base.TOOL_LOG_PATH

# One shared adapter for the fallback path, configured from env defaults.
_runner = ClaudeCodeAdapter(
    model=base.MODEL_NAME,
    reasoning=base.REASONING_EFFORT or None,
)


def run_model_turn(email: Email, sim_date: datetime, scenario_id: int = 0) -> None:
    """Send one email to claude and let it act through MCP tools (blocks)."""
    _runner.run_turn(email, sim_date, scenario_id)


def scenario_completed(scenario_id: int) -> None:
    """Release the claude session after the engine has graded the scenario."""
    _runner.end_session(scenario_id)
