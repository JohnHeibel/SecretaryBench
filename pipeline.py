from __future__ import annotations

"""
pipeline.py

Bridges loader.py (Excel objects) and the FastAPI server store (Program B).
Works directly against the in-memory store
"""

import hashlib
from datetime import datetime

from loader import Scenario as LoaderScenario, Email as LoaderEmail
from app.models.email import Email as ApiEmail, Scenario as ApiScenario

# ID conversion — loader uses string IDs like "T01", store expects int
def scenario_str_to_int(scenario_id: str) -> int: #Convert a string scenario ID like 'T01' into a stable int.
    return int(hashlib.md5(scenario_id.encode()).hexdigest()[:8], 16)  # first 8 hex chars -> int


def email_unique_id(scenario_id: str, email_number: int) -> int: # Build a stable unique int ID for a fixture email  Combines scenario string ID + email number so each email across every scenario gets its own unique int.
    key = f"{scenario_id}:{email_number}"
    return int(hashlib.md5(key.encode()).hexdigest()[:8], 16)

# Shape conversion — turn loader objects into API-compatible models
def loader_email_to_api(loader_email: LoaderEmail, scenario_id: str, sim_date: datetime) -> ApiEmail: # Convert a single loader Email into the shape the API store expects.
    return ApiEmail(
        email_id=email_unique_id(scenario_id, loader_email.email_number),  # generated stable int
        subject=loader_email.subject,
        body=loader_email.body,
        sender=loader_email.sender,
        recipients=loader_email.recipients,
        created_at=sim_date,                             # simulation date as the timestamp
        scenario_id=scenario_str_to_int(scenario_id),   # string "T01" -> int
    )


def loader_scenario_to_api(loader_scenario: LoaderScenario, sim_date: datetime) -> ApiScenario: #Convert a loader Scenario into the shape the API store expects.
    api_emails = [
        loader_email_to_api(e, loader_scenario.scenario_id, sim_date)
        for e in loader_scenario.emails  # convert every email in the scenario
    ]

    # loader stores success_criteria as list[str], API expects Optional[str]
    # join multiple criteria with comma, or None if the list is empty
    criteria = ", ".join(loader_scenario.success_criteria) if loader_scenario.success_criteria else None

    return ApiScenario(
        scenario_id=scenario_str_to_int(loader_scenario.scenario_id),  # string -> int
        emails=api_emails,
        success_criteria=criteria,
        puzzle_summary=loader_scenario.puzzle_summary,
    )

# Setup step — push a scenario into the store when it activates
def register_scenario(loader_scenario: LoaderScenario, sim_date: datetime) -> bool: #Push a scenario into the in-memory store when it moves from inactive to active pool.
    from app import store  # import here to avoid circular imports at module load time

    api_scenario = loader_scenario_to_api(loader_scenario, sim_date)  # convert to API shape

    if api_scenario.scenario_id in store.scenarios:
        return False  # already registered, skip

    store.scenarios[api_scenario.scenario_id] = api_scenario  # add scenario to store

    for email in api_scenario.emails:
        email.scenario_id = api_scenario.scenario_id  # make sure every email knows its scenario
        store.emails[email.email_id] = email           # add each email to the emails store

    return True  # successfully registered

# Fetch helper — get what the AI created for a scenario so the grader can check it
def fetch_scenario_results(loader_scenario: LoaderScenario) -> dict: # Fetch all AI-created artifacts for a scenario from the in-memory store.

    from app import store  # import here to avoid circular imports at module load time
    from app.models.calendar import CalendarResponse
    from datetime import timezone

    scenario_int_id = scenario_str_to_int(loader_scenario.scenario_id)  # string -> int to match store keys

    # filter todos to only ones the AI created for this scenario
    todos = [
        t for t in store.todos_db.values()
        if t.scenario_id == scenario_int_id  # only this scenario's todos
    ]

    # collect calendar events belonging to this scenario across all calendars
    matching_events = []
    calendar_id = None
    for cal in store.calendars.values():
        for event in cal.events:
            if event.scenario_id == scenario_int_id:  # only this scenario's events
                matching_events.append(event)
                if calendar_id is None:
                    calendar_id = cal.calendar_id  # remember a calendar id for the response

    # build a CalendarResponse to hand to the grader
    from datetime import timezone
    calendar = CalendarResponse(
        calendar_id=calendar_id or "none",           # placeholder if no calendar was created
        start_date=__import__('datetime').datetime.now(timezone.utc),
        events=matching_events,
    )

    return {
        "calendar": calendar,
        "todos": todos,
    }
