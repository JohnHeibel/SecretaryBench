"""
harness.claude_p — the Claude Code (`claude -p`) adapter.

Everything harness-agnostic lives in harness.base / harness.cli_base; this file
holds only what is specific to the claude CLI: its flag surface and the
OpenRouter / custom-base-URL env wiring.

Session continuity: the system prompt is injected via --append-system-prompt on
email 1 only; emails 2..N use --resume <session_id> (scraped from the stream-json
system/init event), which carries the prompt and prior turns forward.
"""
from __future__ import annotations

import os
from typing import Optional

from harness import base
from harness.cli_base import CLISubprocessAdapter


class ClaudeCodeAdapter(CLISubprocessAdapter):
    """Drives Claude Code non-interactively as the benchmark harness.

    Each scenario maps to one claude session. Emails 2..N resume it so
    conversation accumulates and compaction can fire automatically. Supports
    OpenRouter / a custom endpoint via ANTHROPIC_BASE_URL + ANTHROPIC_API_KEY.
    """

    binary = "claude"

    def build_command(self, user_msg: str, existing_session: Optional[str]) -> list[str]:
        cmd = [
            "claude", "-p",
            "--output-format", "stream-json",
            "--verbose",
            "--tools", "",
            "--permission-mode", "bypassPermissions",
            "--mcp-config", base.MCP_CONFIG,
            "--strict-mcp-config",
            "--disallowed-tools", base.DISALLOWED_TOOLS,
        ]

        if existing_session is None:
            cmd += ["--model", self.model, "--append-system-prompt", self.system_prompt]
        else:
            cmd += ["--resume", existing_session]

        if self.reasoning:
            cmd += ["--effort", self.reasoning]

        cmd.append(user_msg)
        return cmd

    def build_env(self) -> Optional[dict[str, str]]:
        if not self.api_base and not self.openrouter_key:
            return None  # inherit the parent env unchanged
        env = os.environ.copy()
        if self.api_base:
            env["ANTHROPIC_BASE_URL"] = self.api_base
        if self.openrouter_key:
            env["ANTHROPIC_API_KEY"] = self.openrouter_key
        return env
