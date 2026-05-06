from __future__ import annotations

"""
pipeline.py

Bridges loader.py (Excel objects) to the FastAPI server store via HTTP.

WHY HTTP, NOT DIRECT MUTATION:
This file used to mutate `app.store` directly (`store.scenarios[id] = ...`).
That only works when engine.py and FastAPI live in the same Python process.
They don't — uvicorn runs in its own process, so direct mutation just edits
engine's local copy of the module while FastAPI's store stays empty. The
result was that every model action 404'd because the scenario it was tagged
with didn't exist in FastAPI's view of the world. All reads/writes now go
through HTTP so engine, model_runner (via MCP), and the grader all see one
consistent store.
"""

import hashlib
import os
from datetime import datetime, timezone

import httpx

from loader import Scenario as LoaderScenario, Email as LoaderEmail
from app.models.email import Email as ApiEmail, Scenario as ApiScenario
from app.models.todo import TodoResponse
from app.models.calendar import CalendarResponse, EventResponse


API_BASE = os.environ.get("AISA_API_BASE_URL", "http://localhost:8000")
_TIMEOUT = 10.0


# ID conversion — loader uses string IDs like "T01", store expects int
def scenario_str_to_int(scenario_id: str) -> int:  # Convert a string scenario ID like 'T01' into a stable int.
    return int(hashlib.md5(scenario_id.encode()).hexdigest()[:8], 16)


def email_unique_id(scenario_id: str, email_number: int) -> int:  # Stable unique int ID per fixture email.
    key = f"{scenario_id}:{email_number}"
    return int(hashlib.md5(key.encode()).hexdigest()[:8], 16)


# Shape conversion — turn loader objects into API-compatible models
def loader_email_to_api(loader_email: LoaderEmail, scenario_id: str, sim_date: datetime) -> ApiEmail:
    return ApiEmail(
        email_id=email_unique_id(scenario_id, loader_email.email_number),
        subject=loader_email.subject,
        body=loader_email.body,
        sender=loader_email.sender,
        recipients=loader_email.recipients,
        created_at=sim_date,
        scenario_id=scenario_str_to_int(scenario_id),
    )


def loader_scenario_to_api(loader_scenario: LoaderScenario, sim_date: datetime) -> ApiScenario:
    api_emails = [
        loader_email_to_api(e, loader_scenario.scenario_id, sim_date)
        for e in loader_scenario.emails
    ]
    criteria = ", ".join(loader_scenario.success_criteria) if loader_scenario.success_criteria else None
    return ApiScenario(
        scenario_id=scenario_str_to_int(loader_scenario.scenario_id),
        emails=api_emails,
        success_criteria=criteria,
        puzzle_summary=loader_scenario.puzzle_summary,
    )


# Setup step — push a scenario into the store via HTTP
def register_scenario(loader_scenario: LoaderScenario, sim_date: datetime) -> bool:
    """POST a scenario (with embedded fixture emails) to the FastAPI store.

    Returns True on creation. Returns False if the scenario was already
    registered (FastAPI returns 409). Any other non-2xx raises.
    """
    api_scenario = loader_scenario_to_api(loader_scenario, sim_date)
    payload = api_scenario.model_dump(mode="json")  # serializes datetimes to ISO 8601
    resp = httpx.post(f"{API_BASE}/scenarios/", json=payload, timeout=_TIMEOUT)
    if resp.status_code == 409:
        return False
    resp.raise_for_status()
    return True


# Fetch helper — read AI-created artifacts for a scenario from the FastAPI store
def fetch_scenario_results(loader_scenario: LoaderScenario) -> dict:
    """Read every todo and event the model created for one scenario.

    Returns a dict with the keys the grader expects:
      - "calendar": a CalendarResponse holding ONLY this scenario's events
      - "todos":    list[TodoResponse] for this scenario
    """
    scenario_int_id = scenario_str_to_int(loader_scenario.scenario_id)

    # Pull all todos and filter client-side (FastAPI doesn't have a per-scenario filter).
    todos_resp = httpx.get(f"{API_BASE}/todos/", timeout=_TIMEOUT)
    todos_resp.raise_for_status()
    todos = [TodoResponse(**t) for t in todos_resp.json() if t.get("scenario_id") == scenario_int_id]

    # Pull all calendars (each carries its events embedded) and collect matches.
    cals_resp = httpx.get(f"{API_BASE}/calendars/", timeout=_TIMEOUT)
    cals_resp.raise_for_status()
    cals_list = cals_resp.json()

    matching_events: list[EventResponse] = []
    chosen_calendar_id: str | None = None
    for cal in cals_list:
        for event_dict in cal.get("events", []):
            if event_dict.get("scenario_id") == scenario_int_id:
                matching_events.append(EventResponse(**event_dict))
                if chosen_calendar_id is None:
                    chosen_calendar_id = cal["calendar_id"]

    calendar = CalendarResponse(
        calendar_id=chosen_calendar_id or "none",
        start_date=datetime.now(timezone.utc),
        events=matching_events,
    )

    return {
        "calendar": calendar,
        "todos": todos,
    }
