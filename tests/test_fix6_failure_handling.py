"""FIX-6: a harness crash/timeout is recorded, graded (not skipped), and the
session is reset so the chain continues with a fresh session."""
from __future__ import annotations

from datetime import datetime, timezone

import httpx
import pytest

from harness import base
from harness.cli_base import CLISubprocessAdapter, HarnessAdapter
from loader import Email, Scenario


def _server_up() -> bool:
    try:
        return httpx.get("http://localhost:8000/", timeout=2).status_code == 200
    except Exception:
        return False


# --- adapter session reset on failure (offline) ----------------------------

def test_run_turn_resets_session_on_failure(monkeypatch):
    monkeypatch.setattr(base, "bootstrap_calendar", lambda sim_date: "cal-x")
    monkeypatch.setattr(base, "TOKEN_LOG_PATH", "")
    monkeypatch.setattr(base, "TOOL_LOG_PATH", "")

    class CrashAdapter(CLISubprocessAdapter):
        binary = "false"

        def build_command(self, user_msg, existing_session):
            return ["false"]  # exits 1 immediately, no output -> RuntimeError

    a = CrashAdapter(conversation_continuity=True)
    a.start_session(5)
    a._sessions[5] = "stale-session-id"  # pretend a prior turn left a session

    email = Email(1, "s", "b", "x", ["y"], None)
    with pytest.raises(RuntimeError):
        a.run_turn(email, datetime(2000, 1, 1, tzinfo=timezone.utc), 5)

    # The dead session must be cleared so the next email starts fresh.
    assert a._sessions[5] is None


def test_successful_turn_keeps_session(monkeypatch):
    monkeypatch.setattr(base, "bootstrap_calendar", lambda sim_date: "cal-x")
    monkeypatch.setattr(base, "TOKEN_LOG_PATH", "")
    monkeypatch.setattr(base, "TOOL_LOG_PATH", "")

    class FakeClaude(CLISubprocessAdapter):
        binary = "true"

        def build_command(self, user_msg, existing_session):
            # `printf` a stream-json init line, then succeed.
            return ["printf", '{"type":"system","subtype":"init","session_id":"sess-9"}\\n']

    a = FakeClaude(conversation_continuity=True)
    a.start_session(7)
    a.run_turn(Email(1, "s", "b", "x", ["y"], None),
               datetime(2000, 1, 1, tzinfo=timezone.utc), 7)
    assert a._sessions[7] == "sess-9"


# --- engine records the failure and the chain continues (needs server) -----

@pytest.mark.skipif(not _server_up(), reason="needs uvicorn on :8000")
def test_engine_records_failure_continues_and_grades(monkeypatch, tmp_path):
    import harness
    import bench_logger
    import engine

    monkeypatch.setattr(base, "bootstrap_calendar", lambda sim_date: "cal-x")
    monkeypatch.setattr(bench_logger, "DELIVERY_LOG_PATH", str(tmp_path / "del.jsonl"))

    class AlwaysCrash(HarnessAdapter):
        def __init__(self, **kw):
            self._s: dict = {}

        def start_session(self, sid):
            self._s[sid] = None

        def run_turn(self, email, sim_date, scenario_id):
            raise RuntimeError("boom: simulated bad --flag")

        def resume_session(self, sid):
            pass

        def end_session(self, sid):
            self._s.pop(sid, None)

    monkeypatch.setitem(harness.HARNESS_REGISTRY, "crash", AlwaysCrash)

    scen = Scenario(
        scenario_id="CRASH1", scenario_type="X",
        emails=[Email(1, "s1", "b1", "a", ["b"], "CC-{date}"),
                Email(2, "s2", "b2", "a", ["b"], "TC-{date}")],
        success_criteria=["CC-{date}", "TC-{date}"], puzzle_summary=None,
    )

    result = engine.run_simulation(scenarios=[scen], sim_days=8, seed=0,
                                   verbose=False, harness="crash")

    # The chain completed despite both turns crashing (it did not break).
    assert result["remaining_active"] == 0, result
    # It was still graded — nothing was created, so it scores 0 of a real max.
    assert result["total_max"] >= 1, result
    assert result["total_score"] == 0, result

    # Both crashed turns were recorded as failures in the delivery log.
    import json
    rows = [json.loads(l) for l in open(tmp_path / "del.jsonl") if l.strip()]
    crash_rows = [r for r in rows if r.get("scenario_id") == "CRASH1"]
    assert len(crash_rows) == 2, crash_rows
    assert all(r["status"] == "failed" for r in crash_rows), crash_rows
    assert all("boom" in (r["error"] or "") for r in crash_rows), crash_rows
