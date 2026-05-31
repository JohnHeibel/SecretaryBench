"""
sb.engine — the serve/grade loop.

Walks a precomputed Plan; for each email it renders the body, lets a `model`
act against a Store, snapshots the per-turn delta, and grades against the
email's node state. The model is injected so tests use a deterministic mock and
the real run plugs in the harness-backed model (phase 2).

The Store here is in-memory (and also the shape the grader reads). For a live run
the same Obj lists are built by snapshotting the FastAPI store, filtered to a node
via the email_id attribution tag.
"""
from __future__ import annotations

import itertools
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import Callable, Optional, Protocol

from sb import resolver
from sb.grader import EmailResult, NodeState, Obj, TurnDelta, grade_email
from sb.resolver import Context
from sb.schema import Corpus, Email
from sb.scheduler import Plan


@dataclass
class _Record:
    obj: Obj
    node: str
    oid: int


class Store:
    """Minimal in-memory calendar/todo store with email_id attribution.

    Mirrors what the model can do via MCP (create/update/delete events & todos)
    and exposes node-scoped state + per-turn deltas for grading.
    """

    def __init__(self, corpus: Corpus):
        self._corpus = corpus
        self._records: list[_Record] = []
        self._ids = itertools.count(1)

    # --- model-facing mutations (email_id is the active email being served) ---

    def create_event(self, email_id: str, title: str, start: datetime,
                      end: Optional[datetime] = None, description: str = "") -> int:
        return self._add("event", email_id, title, start, end, description)

    def create_todo(self, email_id: str, title: str, due: datetime,
                    description: str = "") -> int:
        return self._add("todo", email_id, title, due, None, description)

    def update_event(self, oid: int, *, start: Optional[datetime] = None,
                     end: Optional[datetime] = None, title: Optional[str] = None) -> None:
        rec = self._by_id(oid)
        if start is not None:
            rec.obj.when = start
        if end is not None:
            rec.obj.end = end
        if title is not None:
            rec.obj.title = title

    def delete(self, oid: int) -> None:
        self._records = [r for r in self._records if r.oid != oid]

    def find_in_node(self, node: str, kind: str, keyword: str) -> Optional[int]:
        kw = keyword.lower()
        for r in self._records:
            if r.node == node and r.obj.kind == kind and kw in r.obj.title.lower():
                return r.oid
        return None

    # --- grading views ---

    def node_state(self, node: str) -> NodeState:
        ev = [r.obj for r in self._records if r.node == node and r.obj.kind == "event"]
        td = [r.obj for r in self._records if r.node == node and r.obj.kind == "todo"]
        return NodeState(events=ev, todos=td)

    def snapshot_ids(self) -> set[int]:
        return {r.oid for r in self._records}

    def delta_since(self, before: set[int]) -> TurnDelta:
        new = [r for r in self._records if r.oid not in before]
        return TurnDelta(
            events=[r.obj for r in new if r.obj.kind == "event"],
            todos=[r.obj for r in new if r.obj.kind == "todo"],
        )

    # --- internals ---

    def _add(self, kind, email_id, title, when, end, description) -> int:
        oid = next(self._ids)
        node = self._corpus.emails[email_id].node
        self._records.append(_Record(
            obj=Obj(kind=kind, title=title, when=when, email_id=email_id,
                    description=description, end=end),
            node=node, oid=oid,
        ))
        return oid

    def _by_id(self, oid: int) -> _Record:
        for r in self._records:
            if r.oid == oid:
                return r
        raise KeyError(oid)


class Model(Protocol):
    """A model acts on one served email by mutating the store."""
    def __call__(self, email: Email, rendered_body: str, ctx: Context, store: Store) -> None: ...


@dataclass
class RunResult:
    results: dict[str, EmailResult] = field(default_factory=dict)

    @property
    def passed(self) -> int:
        return sum(1 for r in self.results.values() if r.passed)

    @property
    def total(self) -> int:
        return len(self.results)

    def score(self) -> float:
        return self.passed / self.total if self.total else 0.0


def run(corpus: Corpus, plan: Plan, model: Model, store: Optional[Store] = None) -> RunResult:
    store = store or Store(corpus)
    out = RunResult()

    for day, batch in enumerate(plan.per_day):
        for email_id in batch:
            email = corpus.emails[email_id]
            ctx = Context(serve=plan.serve_date[email_id], anchors=plan.anchors, facts=corpus.facts)
            rendered = resolver.render_body(email.body, ctx).text

            before = store.snapshot_ids()
            model(email, rendered, ctx, store)
            turn = store.delta_since(before)

            result = grade_email(email.answer, ctx, store.node_state(email.node), turn)
            out.results[email_id] = result

    return out
