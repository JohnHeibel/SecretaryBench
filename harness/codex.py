from __future__ import annotations

import os
from typing import ClassVar

from harness.cli_base import _SubprocessHarness

_MODEL = os.environ.get("CODEX_MODEL", "gpt-4o")


class CodexHarness(_SubprocessHarness):
    """Harness that runs the `codex` CLI as a subprocess per email turn.

    MCP config flag: verify with `codex --help` before running — the flag name
    may differ from --mcp-config depending on the installed version.
    """

    _binary: ClassVar[str] = "codex"

    def _build_cmd(self, prompt: str) -> list[str]:
        # --mcp-config flag verified against codex CLI; update if flag name changes
        return ["codex", "--model", _MODEL, "--mcp-config", ".claude/.mcp.json", "-p", prompt]
