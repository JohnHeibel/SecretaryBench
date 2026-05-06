from __future__ import annotations

import uuid
from datetime import datetime, timedelta, timezone

from app import store
from app.models.todo import TodoResponse
from app.models.calendar import CalendarResponse, EventResponse
from loader import Email, Scenario
from pipeline import scenario_str_to_int


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


def _clear_store():
    store.todos_db.clear()
    store.calendars.clear()
    store.scenarios.clear()
    store.emails.clear()


def stub_perfect_model(email: Email, sim_date: datetime) -> None:
    """A model that always does the correct thing based on the criteria."""
    criteria = email.success_criteria or ""

    scenario_id = None
    for s in SCENARIOS:
        if email in s.emails or email.subject in [e.subject for e in s.emails]:
            scenario_id = scenario_str_to_int(s.scenario_id)
            break
    if scenario_id is None:
        return

    if "TC" in criteria.upper():
        todo_id = str(uuid.uuid4())
        store.todos_db[todo_id] = TodoResponse(
            id=todo_id,
            title=f"Todo from {email.subject}",
            due_date=sim_date + timedelta(days=1),
            created_at=sim_date,
            completed=False,
            scenario_id=scenario_id,
        )

    elif "CC" in criteria.upper():
        cal_id = str(uuid.uuid4())
        event_id = str(uuid.uuid4())
        event = EventResponse(
            event_id=event_id,
            title=f"Meeting from {email.subject}",
            start=sim_date + timedelta(days=2),
            end=sim_date + timedelta(days=2, hours=1),
            scenario_id=scenario_id,
        )
        if cal_id not in store.calendars:
            store.calendars[cal_id] = CalendarResponse(
                calendar_id=cal_id,
                start_date=sim_date,
                events=[event],
            )
        else:
            store.calendars[cal_id].events.append(event)


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
