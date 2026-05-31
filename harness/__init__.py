"""
harness — Harness Abstraction Layer (package form).

The engine talks to this layer; swapping harness or model is one CLI/env value
with zero changes to engine, grader, or MCP. There is exactly ONE shared core
(harness.base) and one live path, so the default `claude -p` run is no longer a
thin, divergent copy of model_runner.

Layout:
    base.py      — single source of truth: MCP config, hidden tools, system
                   prompt, build_user_message/build_prompt, bootstrap_calendar,
                   parse_stream_output + token/tool/compaction logging.
    cli_base.py  — HarnessAdapter interface + CLISubprocessAdapter turn loop.
    claude_p.py  — ClaudeCodeAdapter (claude-specific flags).
    codex.py     — CodexAdapter (stub).

Adding a harness: subclass HarnessAdapter (or CLISubprocessAdapter), register it
in HARNESS_REGISTRY, select it with --harness <name>. See SPRINT5_REMEDIATION §7.
"""
from __future__ import annotations

from datetime import datetime
from typing import Optional

from loader import Email

from harness import base
from harness.base import (  # re-export the shared core for convenience
    STATIC_SYSTEM_PROMPT,
    MCP_CONFIG,
    HIDDEN_TOOLS,
    DISALLOWED_TOOLS,
    build_user_message,
    build_prompt,
    bootstrap_calendar,
    parse_stream_output,
)
from harness.cli_base import HarnessAdapter, CLISubprocessAdapter
from harness.claude_p import ClaudeCodeAdapter
from harness.codex import CodexAdapter


HARNESS_REGISTRY: dict[str, type[HarnessAdapter]] = {
    "claude-code": ClaudeCodeAdapter,
    "codex": CodexAdapter,
}


def get_adapter(harness: str, **kwargs) -> HarnessAdapter:
    """Instantiate the right adapter by name. kwargs flow to __init__."""
    cls = HARNESS_REGISTRY.get(harness)
    if cls is None:
        raise ValueError(f"Unknown harness {harness!r}. Available: {list(HARNESS_REGISTRY.keys())}")
    return cls(**kwargs)


# --- Backwards-compat shims (old harness.py module API) --------------------
# The previous flat module exposed _build_user_message (3-arg, no calendar_id)
# and _extract_session_id. Keep thin aliases so any external smoke-test or
# import keeps working. New code should use harness.base directly.

def _build_user_message(email: Email, sim_date: datetime, scenario_id: int, calendar_id: str = "cal") -> str:
    return build_user_message(email, sim_date, scenario_id, calendar_id)


def _extract_session_id(output: str) -> Optional[str]:
    """Lightweight session-id scrape (no logging side effects).

    Prefer harness.base.parse_stream_output on the live path — it also logs
    tokens/tools and detects compaction. This stays for backward compat.
    """
    import json
    session_id: Optional[str] = None
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


__all__ = [
    "HarnessAdapter",
    "CLISubprocessAdapter",
    "ClaudeCodeAdapter",
    "CodexAdapter",
    "HARNESS_REGISTRY",
    "get_adapter",
    "build_user_message",
    "build_prompt",
    "bootstrap_calendar",
    "parse_stream_output",
    "STATIC_SYSTEM_PROMPT",
    "MCP_CONFIG",
    "HIDDEN_TOOLS",
    "DISALLOWED_TOOLS",
    "base",
]
