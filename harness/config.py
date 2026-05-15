from __future__ import annotations

import os

from harness.base import ModelHarness


def get_harness() -> ModelHarness:
    """Return the harness selected by the HARNESS environment variable.

    Defaults to claude-p if HARNESS is unset.
    Valid values: claude-p | agent-sdk | codex
    """
    name = os.environ.get("HARNESS", "claude-p")
    if name == "claude-p":
        from harness.claude_p import ClaudePHarness
        return ClaudePHarness()
    if name == "agent-sdk":
        from harness.agent_sdk import AgentSDKHarness
        return AgentSDKHarness()
    if name == "codex":
        from harness.codex import CodexHarness
        return CodexHarness()
    raise ValueError(f"Unknown HARNESS={name!r} — valid: claude-p | agent-sdk | codex")
