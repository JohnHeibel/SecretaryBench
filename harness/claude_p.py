from __future__ import annotations

import os
from typing import ClassVar

from harness.cli_base import _SubprocessHarness

_MODEL = os.environ.get("CLAUDE_P_MODEL", "claude-haiku-4-5")


class ClaudePHarness(_SubprocessHarness):
    """Harness that runs `claude -p <prompt>` as a subprocess per email turn."""

    _binary: ClassVar[str] = "claude"

    def _build_cmd(self, prompt: str) -> list[str]:
        return ["claude", "-p", prompt, "--model", _MODEL, "--mcp-config", ".claude/.mcp.json", "--max-turns", "25"]
