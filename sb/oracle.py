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

from datetime import datetime, time

from sb.engine import Store
from sb.resolver import Context, Interval, Value, resolve
from sb.schema import Email


def _target(predicate: dict, ctx: Context) -> Value:
    if "eq" in predicate:
        return resolve(predicate["eq"], ctx)
    if "any_of" in predicate:
        return resolve(predicate["any_of"][0], ctx)
    if "by" in predicate:
        return resolve(predicate["by"], ctx)
    if "in" in predicate:
        win = resolve(predicate["in"], ctx)
        return win.start if isinstance(win, Interval) else win
    raise ValueError(f"oracle cannot satisfy predicate {predicate!r}")


def _as_dt(v: Value, default_hour: int) -> datetime:
    if isinstance(v, datetime):
        return v
    if isinstance(v, Interval):
        v = v.start
    return datetime.combine(v, time(default_hour, 0))


def oracle_model(email: Email, rendered_body: str, ctx: Context, store: Store) -> None:
    ans = email.answer
    if not ans.expect and not ans.forbid:
        return  # correctly take no action

    for entry in ans.expect:
        title = entry.title_match[0] if entry.title_match else "task"

        if entry.action in ("create_event", "reschedule"):
            if entry.count == 0:  # cancellation
                while (oid := store.find_in_node(email.node, "event", title)) is not None:
                    store.delete(oid)
                continue
            start = _as_dt(_target(entry.start, ctx), 9)   # default 9am; only the DAY is graded
            existing = store.find_in_node(email.node, "event", title)
            if existing is not None:                      # reschedule in place
                store.update_event(existing, start=start)
            else:
                store.create_event(email.id, title, start)

        elif entry.action == "create_todo":
            if entry.count == 0:
                while (oid := store.find_in_node(email.node, "todo", title)) is not None:
                    store.delete(oid)
                continue
            due = _as_dt(_target(entry.due or entry.start, ctx), 17)
            store.create_todo(email.id, title, due)
