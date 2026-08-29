"""
sb.oracle — a reference "perfect secretary" that acts straight from the answer
key. Two uses:

  1. Corpus validation: a well-formed corpus must be oracle-solvable (score 1.0).
     If the oracle can't satisfy an answer, the answer key is unsatisfiable.
  2. A control in tests / a worked example of the action contract.

It is NOT a model under test — it reads ground truth. The harness-backed model
(phase 2) reads only the rendered email.
"""
from __future__ import annotations

from datetime import datetime, time, timedelta

from sb.engine import Store
from sb.resolver import Context, Interval, Value, resolve
from sb.schema import Email, Op


def _target(predicate: dict, ctx: Context) -> Value:
    if "eq" in predicate:
        return resolve(predicate["eq"], ctx)
    if "any_of" in predicate:
        return resolve(predicate["any_of"][0], ctx)
    if "by" in predicate:
        return resolve(predicate["by"], ctx)
    if "in" in predicate:
        win = resolve(predicate["in"], ctx)
        if not isinstance(win, Interval):
            return win
        block = resolve(predicate["not_in"], ctx) if "not_in" in predicate else None
        cur = win.start
        while cur <= win.end:
            if not (isinstance(block, Interval) and block.contains(cur)):
                return cur
            cur += timedelta(days=1)
        raise ValueError(f"oracle cannot satisfy fully blocked predicate {predicate!r}")
    raise ValueError(f"oracle cannot satisfy predicate {predicate!r}")


def _as_dt(v: Value, default_hour: int) -> datetime:
    if isinstance(v, datetime):
        return v
    if isinstance(v, Interval):
        v = v.start
    return datetime.combine(v, time(default_hour, 0))


def oracle_model(email: Email, rendered_body: str, ctx: Context, store: Store) -> None:
    for op in email.answer.ops:
        title = op.name.replace("_", " ")     # the grader keys identity on the name (docs/grader-contract.md)

        if op.verb == "cancel":
            while (oid := store.find_in_node(email.node, op.kind, title)) is not None:
                store.delete(oid)
            continue

        when = _as_dt(_target(op.on, ctx), 9 if op.kind == "event" else 17)
        existing = store.find_in_node(email.node, op.kind, title)

        if op.verb == "move":
            # reschedule in place so the obligation stays a single object
            if existing is not None and op.kind == "event":
                store.update_event(existing, start=when)
            elif existing is not None:
                store.delete(existing)
                store.create_todo(email.id, title, when)
            else:
                _create(store, email.id, op.kind, title, when)
        else:  # create
            _create(store, email.id, op.kind, title, when)


def _create(store: Store, email_id: str, kind: str, title: str, when: datetime) -> None:
    if kind == "event":
        store.create_event(email_id, title, when)
    else:
        store.create_todo(email_id, title, when)
