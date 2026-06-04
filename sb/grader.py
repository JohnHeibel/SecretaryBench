"""
sb.grader — the state-based grader.

Grades one email against its answer key by inspecting the calendar/todo STATE
(not the model's edit), so a reschedule done as update-in-place and one done as
delete-then-recreate score identically (ANSWER_KEY_GRAMMAR.md §8).

The answer key is a list of `ops` (create / move / cancel) on named obligations.
Grading reconciles each obligation against the cumulative node state:

  - create / move : exactly ONE object matching the obligation, on the right day.
                    Uniqueness is inherent to an obligation, so a reschedule that
                    leaves a stale duplicate fails (the old `count: 1` job).
  - cancel        : NO object matching the obligation may survive.
  - no ops        : a no-action / FYI / bait email — the turn must create nothing.
                    (Per-email until the day-loop lands; ADR 0001 stage note.)

Inputs per email:
  - answer:   the Answer (ops)
  - ctx:      resolver.Context (this email's serve date + the live anchor table)
  - state:    NodeState — cumulative objects belonging to this email's NODE
  - turn:     TurnDelta — objects CREATED during this email's turn (for no-action)

Output: EmailResult(passed, max=1, details[]). Binary per email.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Optional

from sb import resolver
from sb.resolver import Context, Interval, TimeInterval, Value
from sb.schema import Answer, Op


@dataclass
class Obj:
    """A calendar event or todo in the store (normalized for grading)."""
    kind: str                 # "event" | "todo"
    title: str
    when: datetime            # event.start or todo.due_date
    email_id: str             # attribution tag the model passed
    description: str = ""
    end: Optional[datetime] = None    # event end (for duration checks)


@dataclass
class NodeState:
    events: list[Obj] = field(default_factory=list)
    todos: list[Obj] = field(default_factory=list)


@dataclass
class TurnDelta:
    events: list[Obj] = field(default_factory=list)
    todos: list[Obj] = field(default_factory=list)


@dataclass
class EmailResult:
    passed: bool
    max: int
    details: list[dict] = field(default_factory=list)
    headline: str = ""


def _title_hit(obj: Obj, keywords: list[str]) -> bool:
    hay = f"{obj.title} {obj.description}".lower()
    return all(k.lower() in hay for k in keywords)


def _to_date(v: Value) -> date:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, TimeInterval):
        return v.start.date()
    if isinstance(v, Interval):
        return v.start
    return v


def _matches_value(obj: Obj, expected: Value, tolerance: str) -> bool:
    # Timed expectations compare at clock granularity (the email dictated a time).
    if isinstance(expected, TimeInterval):
        if obj.when != expected.start:
            return False
        # an object that carries an end (an event) must match it too — this is what
        # grades duration; a point object (a todo, end=None) matches on start alone.
        return obj.end is None or obj.end == expected.end
    if isinstance(expected, datetime):
        return obj.when == expected
    if isinstance(expected, Interval):
        return expected.contains(obj.when)
    if tolerance.startswith("within:"):
        n = int(tolerance.split(":", 1)[1].rstrip("d"))
        return abs((obj.when.date() - _to_date(expected)).days) <= n
    # day equality (a bare date grades at whole-day granularity)
    return obj.when.date() == _to_date(expected)


def _predicate_ok(obj: Obj, predicate: Optional[dict], ctx: Context, tolerance: str) -> bool:
    if not predicate:
        return True
    if "eq" in predicate:
        return _matches_value(obj, resolver.resolve(predicate["eq"], ctx), tolerance)
    if "any_of" in predicate:
        return any(_matches_value(obj, resolver.resolve(e, ctx), tolerance)
                   for e in predicate["any_of"])
    if "by" in predicate:
        return ctx.serve <= obj.when.date() <= _to_date(resolver.resolve(predicate["by"], ctx))
    if "in" in predicate:
        window = resolver.resolve(predicate["in"], ctx)
        ok = isinstance(window, Interval) and window.contains(obj.when)
        if ok and "not_in" in predicate:
            block = resolver.resolve(predicate["not_in"], ctx)
            if isinstance(block, Interval) and block.contains(obj.when):
                ok = False
        return ok
    return True


# --- human-readable formatting (so a run explains itself) ------------------

def _h12(dt: datetime) -> str:
    h = dt.hour % 12 or 12
    ap = "AM" if dt.hour < 12 else "PM"
    mm = "" if dt.minute == 0 else f":{dt.minute:02d}"
    return f"{h}{mm} {ap}"


def _fmt_value(v: Value) -> str:
    if isinstance(v, TimeInterval):
        return f"{v.start.strftime('%a %b %d')} {_h12(v.start)}–{_h12(v.end)}"
    if isinstance(v, Interval):
        return f"{v.start.strftime('%a %b %d')}–{v.end.strftime('%b %d')}"
    if isinstance(v, datetime):
        return f"{v.strftime('%a %b %d')} {_h12(v)}"
    return v.strftime("%a %b %d")


def _fmt_obj(o: Obj) -> str:
    when = f"{o.when.strftime('%a %b %d')} {_h12(o.when)}"
    if o.kind == "event" and o.end is not None:
        when += f" ({round((o.end - o.when).total_seconds() / 60)}m)"
    return f'"{o.title}" {when}'


def _describe_predicate(predicate: Optional[dict], ctx: Context) -> str:
    if not predicate:
        return "(any time)"
    if "eq" in predicate:
        return _fmt_value(resolver.resolve(predicate["eq"], ctx))
    if "any_of" in predicate:
        return " or ".join(_fmt_value(resolver.resolve(e, ctx)) for e in predicate["any_of"])
    if "by" in predicate:
        return f"by {_fmt_value(resolver.resolve(predicate['by'], ctx))}"
    if "in" in predicate:
        s = f"within {_fmt_value(resolver.resolve(predicate['in'], ctx))}"
        if "not_in" in predicate:
            s += f" (avoiding {_fmt_value(resolver.resolve(predicate['not_in'], ctx))})"
        return s
    return "(any time)"


def _wrong_when_reason(op: Op, ctx: Context, title_set: list) -> str:
    """A timed obligation that landed on the right DAY but the wrong clock time (or
    the wrong length) should say so, instead of the day-level 'on the wrong day'."""
    exprs = ([op.on["eq"]] if op.on and "eq" in op.on
             else op.on["any_of"] if op.on and "any_of" in op.on else [])
    for e in exprs:
        val = resolver.resolve(e, ctx)
        if isinstance(val, TimeInterval):
            for o in title_set:
                if o.when == val.start and o.end is not None and o.end != val.end:
                    return "the wrong length"
            if any(o.when.date() == val.start.date() for o in title_set):
                return "at the wrong time"
        elif isinstance(val, datetime):
            if any(o.when.date() == val.date() for o in title_set):
                return "at the wrong time"
    return "on the wrong day"


def _grade_op(op: Op, ctx: Context, state: NodeState) -> dict:
    obj_word = "event" if op.kind == "event" else "to-do"
    pool = state.events if op.kind == "event" else state.todos
    title_set = [o for o in pool if _title_hit(o, op.match)]
    kw = "/".join(op.match) or op.name

    if op.verb == "cancel":
        passed = len(title_set) == 0
        actual = "; ".join(_fmt_obj(o) for o in title_set) if title_set else "(nothing — cancelled)"
        reason = "cancelled" if passed else f"should be cancelled, but {len(title_set)} still on the calendar"
        return {"passed": passed, "label": f"cancel ~{kw}",
                "expected": f'{obj_word} ~"{kw}" cancelled', "actual": actual, "reason": reason}

    # create / move: exactly one object matching the obligation, on the right day.
    matched = [o for o in title_set if _predicate_ok(o, op.on, ctx, op.tolerance)]
    count_ok = len(title_set) == 1
    passed = count_ok and len(matched) >= 1

    expected = f'{obj_word} ~"{kw}" @ {_describe_predicate(op.on, ctx)}'
    actual = "; ".join(_fmt_obj(o) for o in title_set) if title_set else "(nothing matching created)"

    if passed:
        reason = "matched"
    elif not title_set:
        reason = f'no {obj_word} titled like "{kw}" was {"moved" if op.verb == "move" else "created"}'
    elif not count_ok:
        reason = f"found {len(title_set)} matching, expected exactly 1 (duplicate / double-booked)"
    elif not matched:
        reason = _wrong_when_reason(op, ctx, title_set)
    else:
        reason = "did not match the expected action"

    return {"passed": passed, "label": f"{op.verb} ~{kw}",
            "expected": expected, "actual": actual, "reason": reason}


def grade_email(answer: Answer, ctx: Context, state: NodeState, turn: TurnDelta) -> EmailResult:
    # No-action email: no ops means "do nothing this turn" (FYI / bait / no scheduling).
    if not answer.ops:
        created = turn.events + turn.todos
        passed = not created
        actual = "; ".join(_fmt_obj(o) for o in created) if created else "(nothing)"
        reason = "correctly took no action" if passed else f"over-acted — created {actual}"
        return EmailResult(passed=passed, max=1, headline=reason, details=[{
            "passed": passed, "label": "no-action",
            "expected": "take no action (FYI / no scheduling needed)",
            "actual": actual, "reason": reason}])

    details = [_grade_op(op, ctx, state) for op in answer.ops]
    passed = all(d["passed"] for d in details)
    if passed:
        headline = "; ".join(d["expected"] for d in details)
    else:
        headline = next(d["reason"] for d in details if not d["passed"])
    return EmailResult(passed=passed, max=1, headline=headline, details=details)
