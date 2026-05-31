"""
harness.cli_base — the HarnessAdapter interface plus a shared base for CLI
subprocess harnesses (claude -p today, codex tomorrow).

CLISubprocessAdapter owns the harness-agnostic turn loop:
  bootstrap calendar -> build user message -> launch subprocess -> block ->
  parse stream output (tokens / tools / compaction / session id) -> store session.

Concrete adapters only implement build_command()/build_env() with their own
flags. This is the "exactly one behavior" guarantee from SPRINT5_REMEDIATION §1.
"""
from __future__ import annotations

import subprocess
from abc import ABC, abstractmethod
from datetime import datetime
from typing import Optional

from loader import Email

from harness import base


class HarnessAdapter(ABC):
    """Interface every harness adapter must implement.

    The engine calls these per scenario:
      1. start_session()   — scenario activates
      2. run_turn()        — one email served (blocks until the harness is done)
      3. resume_session()  — emails 2..N re-attach context (no-op for claude;
                             resume is handled inside run_turn via --resume)
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

    def scenario_telemetry(self, scenario_id: int) -> dict:
        """Per-scenario telemetry (compaction count + token totals) for FIX-9.

        Non-abstract with a zeroed default so harnesses that emit nothing
        parseable (or stubs) don't have to implement it. Read it before
        end_session(), which releases the underlying state.
        """
        return {"compactions": 0, "input_tokens": 0, "output_tokens": 0, "rounds": 0}


class CLISubprocessAdapter(HarnessAdapter):
    """Shared turn loop for CLI subprocess harnesses.

    Subclasses set `binary` and implement `build_command()` (and optionally
    `build_env()`). Everything else — calendar bootstrap, the user message,
    timeout handling, stream-json parsing + logging, and session continuity —
    comes from harness.base so there is one source of truth.
    """

    binary: str = "claude"

    def __init__(
        self,
        model: str = base.MODEL_NAME,
        reasoning: Optional[str] = None,
        api_base: Optional[str] = None,
        openrouter_key: Optional[str] = None,
        system_prompt: Optional[str] = None,
        per_turn_timeout: float = base.PER_TURN_TIMEOUT_S,
        conversation_continuity: Optional[bool] = None,
    ):
        self.model = model
        self.reasoning = reasoning
        self.api_base = api_base
        self.openrouter_key = openrouter_key
        self.system_prompt = system_prompt or base.STATIC_SYSTEM_PROMPT
        self.per_turn_timeout = per_turn_timeout
        # Default comes from the env knob (FIX-8); the engine may override it.
        self.conversation_continuity = (
            base.CONVERSATION_CONTINUITY if conversation_continuity is None else conversation_continuity
        )
        self._sessions: dict[int, Optional[str]] = {}

    # --- Subclass hooks ----------------------------------------------------

    def build_command(self, user_msg: str, existing_session: Optional[str]) -> list[str]:
        """Return the argv for one turn. existing_session is None on email 1."""
        raise NotImplementedError

    def build_env(self) -> Optional[dict[str, str]]:
        """Return the subprocess env (or None to inherit the parent env)."""
        return None

    # --- HarnessAdapter ----------------------------------------------------

    def start_session(self, scenario_id: int) -> None:
        self._sessions[scenario_id] = None

    def run_turn(self, email: Email, sim_date: datetime, scenario_id: int) -> None:
        calendar_id = base.bootstrap_calendar(sim_date)
        user_msg = base.build_user_message(email, sim_date, scenario_id, calendar_id)
        email_index = base.next_email_index(scenario_id)
        existing_session = self._sessions.get(scenario_id) if self.conversation_continuity else None

        cmd = self.build_command(user_msg, existing_session)
        env = self.build_env()

        try:
            proc = subprocess.Popen(
                cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, env=env
            )
            try:
                stdout, stderr = proc.communicate(timeout=self.per_turn_timeout)
            except subprocess.TimeoutExpired:
                proc.kill()
                proc.communicate()
                raise TimeoutError(
                    f"Per-turn {self.per_turn_timeout:.0f}s timeout on scenario {scenario_id}."
                )

            if proc.returncode != 0:
                raise RuntimeError(
                    f"{self.binary} exited {proc.returncode} on scenario {scenario_id}: {stderr.strip()}"
                )

            session_id = base.parse_stream_output(stdout, scenario_id, email_index)
            if self.conversation_continuity and session_id:
                self._sessions[scenario_id] = session_id
            base.note_turn_success()

        except Exception as exc:
            base.note_turn_failure(scenario_id, email_index, exc)
            # FIX-6: a failed turn must not leave a stale/dead session for the
            # next email in the chain to --resume. Clear it so email N+1 starts
            # fresh instead of resuming a session that never completed.
            self._sessions[scenario_id] = None
            raise

    def resume_session(self, scenario_id: int) -> None:
        pass  # handled inside run_turn via --resume

    def scenario_telemetry(self, scenario_id: int) -> dict:
        return base.scenario_telemetry(scenario_id)

    def end_session(self, scenario_id: int) -> None:
        self._sessions.pop(scenario_id, None)
        base.release_scenario(scenario_id)
