from datetime import datetime
from typing import Optional
from pydantic import BaseModel


class EventCreate(BaseModel):
    # what the agent sends. scenario_id is required so every event is attributable.
    title: str
    description: Optional[str] = None
    start: datetime
    end: datetime
    scenario_id: int


class EventResponse(BaseModel):
    event_id: str
    title: str
    description: Optional[str] = None
    start: datetime
    end: datetime
    # Always set on creation (EventCreate requires it), so the response is non-optional (FIX-13).
    scenario_id: int


class CalendarCreate(BaseModel):
    start_date: datetime


class CalendarResponse(BaseModel):
    calendar_id: str
    start_date: datetime
    events: list[EventResponse] = []
