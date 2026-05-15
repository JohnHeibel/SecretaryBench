from __future__ import annotations

import logging
import subprocess
import time
from datetime import datetime
from typing import ClassVar

from loader import Email
from harness.base import TurnResult, build_prompt, bootstrap_calendar, PER_TURN_TIMEOUT_S

logger = logging.getLogger(__name__)


class _SubprocessHarness:
    """Shared base for CLI subprocess harnesses (claude -p, codex).

    Subclasses set _binary and implement _build_cmd().
    """

    _binary: ClassVar[str]

    def __init__(self) -> None:
        self._calendar_id: str | None = None
        self._sessions: dict[int, str] = {}

    def _build_cmd(self, prompt: str) -> list[str]:
        raise NotImplementedError

    def run_turn(self, email: Email, sim_date: datetime, scenario_id: int) -> TurnResult | None:
        self._calendar_id = bootstrap_calendar(sim_date, self._calendar_id)
        prompt = build_prompt(email, sim_date, scenario_id, self._calendar_id)
        cmd = self._build_cmd(prompt)
        t0 = time.monotonic()
        try:
            proc = subprocess.run(cmd, capture_output=True, text=True, timeout=PER_TURN_TIMEOUT_S)
        except subprocess.TimeoutExpired:
            logger.error("scenario %d: %s timed out after %.0fs", scenario_id, self._binary, PER_TURN_TIMEOUT_S)
            raise TimeoutError(f"scenario {scenario_id}: {self._binary} exceeded {PER_TURN_TIMEOUT_S}s")
        elapsed = time.monotonic() - t0
        if proc.returncode != 0:
            logger.error("scenario %d: %s exited %d\nstderr: %s", scenario_id, self._binary, proc.returncode, proc.stderr[:500])
        return TurnResult(scenario_id=scenario_id, elapsed_s=elapsed)

    def scenario_completed(self, scenario_id: int) -> None:
        self._sessions.pop(scenario_id, None)

    def shutdown(self) -> None:
        pass
