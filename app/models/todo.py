from datetime import datetime
from typing import Optional
from pydantic import BaseModel


class TodoCreate(BaseModel):
    # what the agent sends. scenario_id is required so every todo is attributable.
    title: str
    description: Optional[str] = None
    due_date: datetime
    scenario_id: int
    calendar_event_id: Optional[str] = None


class TodoUpdate(BaseModel):
    title: Optional[str] = None
    description: Optional[str] = None
    due_date: Optional[datetime] = None
    completed: Optional[bool] = None


class TodoResponse(BaseModel):
    id: str
    title: str
    description: Optional[str] = None
    due_date: datetime
    created_at: datetime
    completed: bool = False
    # Always set on creation (TodoCreate requires it), so the response is non-optional (FIX-13).
    scenario_id: int
    calendar_event_id: Optional[str] = None #need linking to where the calendar event is. Available to change.