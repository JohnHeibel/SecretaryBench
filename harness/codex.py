"""
harness.codex — stub for the OpenAI Codex CLI harness.

Forward-looking only. The MCP tool surface and prompt come from harness.base
(the fairness invariant in SPRINT5_REMEDIATION §7.4); a real implementation
just needs Codex's own launch flags and resume mechanism. See §7.3 for the
per-harness recipe and §7.6 (FIX-14) for proving the abstraction with a second
real harness.
"""
from __future__ import annotations

from datetime import datetime
from typing import Optional

from sb.schema import Email
from harness.cli_base import HarnessAdapter


class CodexAdapter(HarnessAdapter):
    """Stub for the OpenAI Codex harness. MCP config identical to ClaudeCodeAdapter."""

    def __init__(self, model: str = "o4-mini", api_key: Optional[str] = None, **kwargs):
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
