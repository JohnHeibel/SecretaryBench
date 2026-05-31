"""
sb.live.mcp_app — the model-facing MCP tool surface for a live run.

Spawned fresh by each `claude -p` call; routes every tool call over HTTP to the
persistent store (sb.live.store_app). The model sees ONE inbox/context for the
whole run and a powerful search_inbox for facts that have scrolled out.

    python -m sb.live.mcp_app          # stdio MCP server (claude --mcp-config)
"""
from __future__ import annotations

import os
from typing import Any, Optional

import httpx
from mcp.server.fastmcp import FastMCP

STORE = os.environ.get("STORE_BASE_URL", "http://localhost:8077")
_client = httpx.Client(base_url=STORE, timeout=15.0)
mcp = FastMCP("secretary")


def _call(method: str, path: str, **kw: Any) -> Any:
    r = _client.request(method, path, **kw)
    if r.status_code >= 400:
        raise RuntimeError(f"HTTP {r.status_code} {method} {path}: {r.text}")
    return None if r.status_code == 204 or not r.content else r.json()


# --- calendar events ---

@mcp.tool()
def create_event(email_id: str, title: str, start: str, end: str, description: str = "") -> dict:
    """Create a calendar event. email_id = the id of the email you are acting on (REQUIRED,
    given to you in the email header). start/end are ISO 8601 (e.g. '2026-06-08T14:00:00')."""
    return _call("POST", "/events", json={"email_id": email_id, "title": title,
                                          "start": start, "end": end, "description": description})


@mcp.tool()
def update_event(event_id: str, title: str | None = None, start: str | None = None,
                 end: str | None = None, description: str | None = None) -> dict:
    """Modify an existing event (e.g. to reschedule). Only the fields you pass change."""
    body = {k: v for k, v in dict(title=title, start=start, end=end, description=description).items()
            if v is not None}
    return _call("PATCH", f"/events/{event_id}", json=body)


@mcp.tool()
def delete_event(event_id: str) -> dict:
    """Delete an event."""
    return _call("DELETE", f"/events/{event_id}")


@mcp.tool()
def list_events() -> list[dict]:
    """List all calendar events you have created."""
    return _call("GET", "/events")


# --- todos ---

@mcp.tool()
def create_todo(email_id: str, title: str, due_date: str, description: str = "") -> dict:
    """Create a to-do. email_id = the id of the email you are acting on (REQUIRED).
    due_date is ISO 8601 (e.g. '2026-06-12T17:00:00')."""
    return _call("POST", "/todos", json={"email_id": email_id, "title": title,
                                         "due_date": due_date, "description": description})


@mcp.tool()
def update_todo(todo_id: str, title: str | None = None, due_date: str | None = None,
                description: str | None = None, completed: bool | None = None) -> dict:
    """Modify an existing to-do. Only the fields you pass change."""
    body = {k: v for k, v in dict(title=title, due_date=due_date, description=description,
                                  completed=completed).items() if v is not None}
    return _call("PATCH", f"/todos/{todo_id}", json=body)


@mcp.tool()
def delete_todo(todo_id: str) -> dict:
    """Delete a to-do."""
    return _call("DELETE", f"/todos/{todo_id}")


@mcp.tool()
def list_todos() -> list[dict]:
    """List all to-dos you have created."""
    return _call("GET", "/todos")


# --- inbox retrieval (the whole point: recover facts from earlier emails) ---

@mcp.tool()
def search_inbox(query: Optional[str] = None, sender: Optional[str] = None,
                 regex: bool = False, recent: Optional[int] = None, limit: int = 10) -> list[dict]:
    """Search your past emails. Use this whenever an email refers to something you don't
    remember (a policy, a date, a person, a prior decision).
      - query: text to match in subject/body/sender (substring, or a regex if regex=true)
      - sender: filter to a sender substring
      - recent: return only the last N emails
      - limit: cap results (default 10)"""
    params: dict[str, Any] = {"regex": regex, "limit": limit}
    if query:
        params["q"] = query
    if sender:
        params["sender"] = sender
    if recent:
        params["recent"] = recent
    return _call("GET", "/inbox/search", params=params)


@mcp.tool()
def get_email(email_id: str) -> dict:
    """Fetch the full text of one past email by its id."""
    return _call("GET", f"/inbox/{email_id}")


if __name__ == "__main__":
    mcp.run()
