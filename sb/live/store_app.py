"""
sb.live.store_app — the persistent in-memory store for a live run.

Runs as a uvicorn process that survives across the many `claude -p` invocations
of a run (each claude call spawns a fresh MCP server subprocess, so state must
live here). Holds the model's calendar events + todos and a searchable inbox of
served emails.

    uvicorn sb.live.store_app:app --port 8077
"""
from __future__ import annotations

import itertools
import re
from typing import Optional

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel

app = FastAPI(title="secretary-bench store")

_ids = itertools.count(1)
_events: dict[str, dict] = {}
_todos: dict[str, dict] = {}
_inbox: list[dict] = []         # served emails, in serve order
_warnings: list[dict] = []      # potential gaming / instrumentation events, read by the runner
_today: Optional[str] = None    # the current simulated day; list_new_emails filters to it


# --- payloads ---

class EventIn(BaseModel):
    email_id: str
    title: str
    start: str
    end: Optional[str] = None
    description: str = ""


class EventPatch(BaseModel):
    title: Optional[str] = None
    start: Optional[str] = None
    end: Optional[str] = None
    description: Optional[str] = None


class TodoIn(BaseModel):
    email_id: str
    title: str
    due_date: str
    description: str = ""


class TodoPatch(BaseModel):
    title: Optional[str] = None
    due_date: Optional[str] = None
    description: Optional[str] = None
    completed: Optional[bool] = None


class InboxIn(BaseModel):
    email_id: str
    sender: str
    recipients: list[str]
    subject: str
    body: str
    served_date: str


class DayIn(BaseModel):
    date: str


def _warn(kind: str, **fields) -> None:
    _warnings.append({"kind": kind, "today": _today, **fields})


def _served_email_ids() -> set[str]:
    return {m["email_id"] for m in _inbox}


def _today_email_ids() -> set[str]:
    return {m["email_id"] for m in _inbox if m["served_date"] == _today}


def _watch_attribution(kind: str, email_id: str, title: str) -> None:
    """Monitor-only for now. The grader trusts email_id attribution; if models start
    hiding work under stale or fake ids, this warning tells us to harden it."""
    if email_id not in _served_email_ids():
        _warn("invalid_email_id", object_kind=kind, email_id=email_id, title=title)
    elif email_id not in _today_email_ids():
        _warn("stale_email_id", object_kind=kind, email_id=email_id, title=title)


# --- lifecycle ---

@app.get("/health")
def health():
    return {"ok": True, "events": len(_events), "todos": len(_todos),
            "inbox": len(_inbox), "today": _today}


@app.post("/reset")
def reset():
    global _ids, _today
    _events.clear()
    _todos.clear()
    _inbox.clear()
    _warnings.clear()
    _ids = itertools.count(1)
    _today = None
    return {"ok": True}


@app.post("/day")
def set_day(d: DayIn):
    """The runner tells the store which simulated day it's currently serving, so
    list_new_emails can return just today's arrivals."""
    global _today
    _today = d.date
    return {"ok": True, "today": _today}


# --- events ---

@app.post("/events", status_code=201)
def create_event(e: EventIn):
    _watch_attribution("event", e.email_id, e.title)
    eid = f"evt_{next(_ids)}"
    _events[eid] = {"id": eid, **e.model_dump()}
    return _events[eid]


@app.get("/events")
def list_events():
    return list(_events.values())


@app.patch("/events/{eid}")
def patch_event(eid: str, p: EventPatch):
    if eid not in _events:
        raise HTTPException(404, f"event {eid} not found")
    _events[eid].update({k: v for k, v in p.model_dump().items() if v is not None})
    return _events[eid]


@app.delete("/events/{eid}")
def delete_event(eid: str):
    if _events.pop(eid, None) is None:
        raise HTTPException(404, f"event {eid} not found")
    return {"deleted": eid}


# --- todos ---

@app.post("/todos", status_code=201)
def create_todo(t: TodoIn):
    _watch_attribution("todo", t.email_id, t.title)
    tid = f"todo_{next(_ids)}"
    _todos[tid] = {"id": tid, **t.model_dump()}
    return _todos[tid]


@app.get("/todos")
def list_todos():
    return list(_todos.values())


@app.patch("/todos/{tid}")
def patch_todo(tid: str, p: TodoPatch):
    if tid not in _todos:
        raise HTTPException(404, f"todo {tid} not found")
    _todos[tid].update({k: v for k, v in p.model_dump().items() if v is not None})
    return _todos[tid]


@app.delete("/todos/{tid}")
def delete_todo(tid: str):
    if _todos.pop(tid, None) is None:
        raise HTTPException(404, f"todo {tid} not found")
    return {"deleted": tid}


# --- inbox (engine posts served emails; model searches them) ---

@app.post("/inbox", status_code=201)
def add_inbox(m: InboxIn):
    _inbox.append(m.model_dump())
    return {"ok": True, "size": len(_inbox)}


@app.get("/inbox/search")
def search_inbox(q: Optional[str] = None, sender: Optional[str] = None,
                 regex: bool = False, recent: Optional[int] = None, limit: int = 10):
    # Monitor-only for now. Search is supposed to be targeted retrieval; if models
    # routinely dump the archive, tighten this endpoint instead of changing scores.
    if not q and not sender:
        _warn("broad_search", limit=limit, recent=recent, inbox_size=len(_inbox))
    if limit == 0 or limit > 25:
        _warn("large_search_result", limit=limit, recent=recent, inbox_size=len(_inbox))
    results = list(_inbox)
    if sender:
        results = [m for m in results if sender.lower() in m["sender"].lower()]
    if q:
        if regex:
            pat = re.compile(q, re.I)
            results = [m for m in results if pat.search(f"{m['subject']}\n{m['body']}\n{m['sender']}")]
        else:
            ql = q.lower()
            results = [m for m in results if ql in f"{m['subject']}\n{m['body']}\n{m['sender']}".lower()]
    if recent:
        results = results[-recent:]
    return results[-limit:] if limit else results


@app.get("/inbox/new")
def list_new_emails():
    """Today's arrivals only — id/sender/subject, no body. Idempotent within the
    day. Defined before /inbox/{email_id} so 'new' isn't swallowed as an id."""
    return [{"email_id": m["email_id"], "sender": m["sender"], "subject": m["subject"],
             "served_date": m["served_date"]}
            for m in _inbox if m["served_date"] == _today]


@app.get("/inbox/{email_id}")
def get_email(email_id: str):
    for m in _inbox:
        if m["email_id"] == email_id:
            return m
    raise HTTPException(404, f"email {email_id} not found")


# --- grading snapshot ---

@app.get("/state")
def state():
    return {"events": list(_events.values()), "todos": list(_todos.values())}


@app.get("/warnings")
def warnings():
    return list(_warnings)
