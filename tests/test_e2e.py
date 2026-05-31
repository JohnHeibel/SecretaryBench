from __future__ import annotations

from datetime import datetime, timedelta

import httpx

from app import store
from loader import Email, Scenario
from pipeline import scenario_str_to_int

# These tests drive the FULL stack: engine -> pipeline (HTTP) -> FastAPI store
# -> grader. They require a live server (`uvicorn app.main:app`) because
# pipeline.register_scenario / fetch_scenario_results talk to it over HTTP.
# The stubs below therefore write artifacts through the API too — writing to the
# in-process app.store would be invisible to the engine's HTTP reads.
API = "http://localhost:8000"


SCENARIOS = [
    Scenario(
        scenario_id="TEST-A",
        scenario_type="T01",
        emails=[
            Email(
                email_number=1,
                subject="Schedule meeting",
                body="Please add a todo for {date+1}",
                sender="boss@company.com",
                recipients=["secretary@company.com"],
                success_criteria="TC-{date+1}",
            )
        ],
        success_criteria=["TC-{date+1}"],
        puzzle_summary="Create a todo",
    ),
    Scenario(
        scenario_id="TEST-B",
        scenario_type="C01",
        emails=[
            Email(
                email_number=1,
                subject="Book conference room",
                body="Add a meeting on {date+2}",
                sender="boss@company.com",
                recipients=["secretary@company.com"],
                success_criteria="CC-{date+2}",
            )
        ],
        success_criteria=["CC-{date+2}"],
        puzzle_summary="Create a calendar event",
    ),
    Scenario(
        scenario_id="TEST-C",
        scenario_type="N01",
        emails=[
            Email(
                email_number=1,
                subject="FYI newsletter",
                body="No action needed on this.",
                sender="info@company.com",
                recipients=["secretary@company.com"],
                success_criteria="No action required",
            )
        ],
        success_criteria=["No action required"],
        puzzle_summary="Do nothing",
    ),
]


def _clear_server() -> None:
    """Wipe the live server's store so each test starts clean.

    Deletes every todo, calendar (cascades its events), and scenario currently
    registered. Safe to call repeatedly; ignores anything already gone.
    """
    for t in httpx.get(f"{API}/todos/", timeout=10).json():
        httpx.delete(f"{API}/todos/{t['id']}", timeout=10)
    for c in httpx.get(f"{API}/calendars/", timeout=10).json():
        httpx.delete(f"{API}/calendars/{c['calendar_id']}", timeout=10)
    for s in httpx.get(f"{API}/scenarios/", timeout=10).json():
        httpx.delete(f"{API}/scenarios/{s['scenario_id']}", timeout=10)


def _clear_store():
    # Local store is unused by the HTTP pipeline, but clear it for good measure.
    store.todos_db.clear()
    store.calendars.clear()
    store.scenarios.clear()
    store.emails.clear()
    _clear_server()


def _scenario_id_for(email: Email) -> int | None:
    for s in SCENARIOS:
        if email in s.emails or email.subject in [e.subject for e in s.emails]:
            return scenario_str_to_int(s.scenario_id)
    return None


def stub_perfect_model(email: Email, sim_date: datetime) -> None:
    """A model that always does the correct thing — via the API, like a real harness."""
    criteria = (email.success_criteria or "").upper()
    scenario_id = _scenario_id_for(email)
    if scenario_id is None:
        return

    if "TC" in criteria:
        httpx.post(f"{API}/todos/", timeout=10, json={
            "title": f"Todo from {email.subject}",
            "due_date": (sim_date + timedelta(days=1)).isoformat(),
            "scenario_id": scenario_id,
        }).raise_for_status()

    elif "CC" in criteria:
        cal = httpx.post(f"{API}/calendars/", timeout=10,
                         json={"start_date": sim_date.isoformat()}).json()
        start = sim_date + timedelta(days=2)
        httpx.post(f"{API}/calendars/{cal['calendar_id']}/events", timeout=10, json={
            "title": f"Meeting from {email.subject}",
            "start": start.isoformat(),
            "end": (start + timedelta(hours=1)).isoformat(),
            "scenario_id": scenario_id,
        }).raise_for_status()


def stub_bad_model(email: Email, sim_date: datetime) -> None:
    """A model that does nothing."""
    pass


def test_perfect_stub_scores_max():
    _clear_store()
    from engine import run_simulation

    result = run_simulation(
        scenarios=list(SCENARIOS),
        sim_days=100,
        seed=42,
        verbose=False,
        model_fn=stub_perfect_model,
    )

    assert result["total_max"] > 0, "Should have criteria to grade"
    assert result["total_score"] == result["total_max"], (
        f"Perfect stub should score max: got {result['total_score']}/{result['total_max']}"
    )


def test_bad_stub_scores_only_no_action():
    _clear_store()
    from engine import run_simulation

    result = run_simulation(
        scenarios=list(SCENARIOS),
        sim_days=100,
        seed=42,
        verbose=False,
        model_fn=stub_bad_model,
    )

    assert result["total_max"] == 3, "Should have 3 criteria total"
    assert result["total_score"] == 1, (
        f"Bad stub should only score the 'No action' scenario: "
        f"got {result['total_score']}/{result['total_max']}"
    )
