"""
sb.grader — the state-based grader.

Grades one email against its answer key by inspecting the calendar/todo STATE
(not the model's edit), so a reschedule done as update-in-place and one done as
delete-then-recreate score identically (ANSWER_KEY_GRAMMAR.md §8).

Inputs per email:
  - answer:   the Answer (expect / forbid / emits)
  - ctx:      resolver.Context (this email's serve date + the live anchor table)
  - state:    NodeState — cumulative objects belonging to this email's NODE
  - turn:     TurnDelta — objects CREATED during this email's turn (for forbid / no-action)

Output: EmailResult(passed, max=1, details[]). Binary per email, with a
"right-action-wrong-time" flag surfaced in details for error analysis (B3).
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Optional

from sb import resolver
from sb.resolver import Context, Interval, Value
from sb.schema import Answer, ExpectEntry, ForbidEntry


@dataclass
class Obj:
    """A calendar event or todo in the store (normalized for grading)."""
    kind: str                 # "event" | "todo"
    title: str
    when: datetime            # event.start or todo.due_date
    email_id: str             # attribution tag the model passed
    description: str = ""
    end: Optional[datetime] = None    # event end (for duration checks)


def _parse_duration_minutes(s: str) -> int:
    """'90m' -> 90, '1h' -> 60, '1h30m' -> 90, '2h' -> 120."""
    import re as _re
    total = 0
    for amount, unit in _re.findall(r"(\d+)\s*([hm])", s.lower()):
        total += int(amount) * (60 if unit == "h" else 1)
    if total == 0:
        raise ValueError(f"bad duration {s!r}")
    return total


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


_EVENT_ACTIONS = {"create_event", "reschedule"}


def _title_hit(obj: Obj, keywords: list[str]) -> bool:
    hay = f"{obj.title} {obj.description}".lower()
    return all(k.lower() in hay for k in keywords)


def _to_date(v: Value) -> date:
    return v.date() if isinstance(v, datetime) else (v.start if isinstance(v, Interval) else v)


def _matches_value(obj_when: datetime, expected: Value, tolerance: str) -> bool:
    if isinstance(expected, Interval):
        return expected.contains(obj_when)
    if tolerance.startswith("within:"):
        n = int(tolerance.split(":", 1)[1].rstrip("d"))
        return abs((obj_when.date() - _to_date(expected)).days) <= n
    if tolerance == "exact_time" and isinstance(expected, datetime):
        return obj_when.replace(second=0, microsecond=0) == expected.replace(second=0, microsecond=0)
    # default: day equality
    return obj_when.date() == _to_date(expected)


def _predicate_ok(obj: Obj, predicate: Optional[dict], ctx: Context, tolerance: str) -> bool:
    if not predicate:
        return True
    if "eq" in predicate:
        return _matches_value(obj.when, resolver.resolve(predicate["eq"], ctx), tolerance)
    if "any_of" in predicate:
        return any(_matches_value(obj.when, resolver.resolve(e, ctx), tolerance)
                   for e in predicate["any_of"])
    if "by" in predicate:
        return obj.when.date() <= _to_date(resolver.resolve(predicate["by"], ctx))
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


def _grade_expect(entry: ExpectEntry, ctx: Context, state: NodeState) -> dict:
    kind = "event" if entry.action in _EVENT_ACTIONS else "to-do"
    pool = state.events if entry.action in _EVENT_ACTIONS else state.todos
    title_set = [o for o in pool if _title_hit(o, entry.title_match)]
    predicate = entry.start if entry.action in _EVENT_ACTIONS else (entry.due or entry.start)

    # duration / count may be literals OR fact references ("@meeting_len"); resolve
    # both against the shared fact table so a policy change propagates here.
    dur_min = (_parse_duration_minutes(resolver.resolve_scalar(entry.duration, ctx.facts))
               if entry.duration else None)
    want_count = (int(resolver.resolve_scalar(str(entry.count), ctx.facts))
                  if entry.count is not None else None)

    def date_ok(o: Obj) -> bool:
        return _predicate_ok(o, predicate, ctx, entry.tolerance)

    def dur_ok(o: Obj) -> bool:
        if dur_min is None or o.end is None:
            return True
        return round((o.end - o.when).total_seconds() / 60) == dur_min

    matched = [o for o in title_set if date_ok(o) and dur_ok(o)]
    need = want_count if want_count is not None else 1
    count_ok = want_count is None or len(title_set) == want_count
    passed = len(matched) >= need and count_ok

    kw = "/".join(entry.title_match) or "any"
    expected = f'{kind} ~"{kw}" @ {_describe_predicate(predicate, ctx)}'
    if dur_min is not None:
        expected += f" · {dur_min} min"
    if want_count is not None:
        expected += f" · exactly {want_count}"

    actual = "; ".join(_fmt_obj(o) for o in title_set) if title_set else "(nothing matching created)"

    if passed:
        reason = "matched"
    elif want_count == 0:
        reason = f"should be cancelled, but {len(title_set)} still on the calendar"
    elif not title_set:
        reason = f'no {kind} titled like "{kw}" was created'
    elif not count_ok:
        extra = " (duplicate / double-booked)" if len(title_set) > (want_count or 0) else ""
        reason = f"created {len(title_set)}, expected exactly {want_count}{extra}"
    elif not any(date_ok(o) for o in title_set):
        reason = f"created, but on the wrong date/time"
    elif dur_min is not None and not any(dur_ok(o) for o in title_set):
        reason = f"right time, wrong length (expected {dur_min} min)"
    else:
        reason = "did not match the expected action"

    return {"passed": passed, "label": f"{kind} ~{kw}",
            "expected": expected, "actual": actual, "reason": reason}


def _grade_forbid(entry: ForbidEntry, turn: TurnDelta) -> dict:
    objs: list[Obj] = []
    if entry.action is None or entry.action in _EVENT_ACTIONS:
        objs += turn.events
    if entry.action is None or entry.action == "create_todo":
        objs += turn.todos
    if entry.title_match:
        objs = [o for o in objs if _title_hit(o, entry.title_match)]
    passed = len(objs) == 0
    what = entry.action or "any action"
    return {"passed": passed, "label": f"forbid {what}",
            "expected": f"do NOT {what}",
            "actual": "; ".join(_fmt_obj(o) for o in objs) if objs else "(nothing)",
            "reason": "respected" if passed else f"created {len(objs)} forbidden object(s)"}


def grade_email(answer: Answer, ctx: Context, state: NodeState, turn: TurnDelta) -> EmailResult:
    # No-action email: empty expect AND empty forbid means "do nothing this turn".
    if not answer.expect and not answer.forbid:
        created = turn.events + turn.todos
        passed = not created
        actual = "; ".join(_fmt_obj(o) for o in created) if created else "(nothing)"
        reason = "correctly took no action" if passed else f"over-acted — created {actual}"
        return EmailResult(passed=passed, max=1, headline=reason, details=[{
            "passed": passed, "label": "no-action",
            "expected": "take no action (FYI / no scheduling needed)",
            "actual": actual, "reason": reason}])

    details = [_grade_expect(e, ctx, state) for e in answer.expect]
    details += [_grade_forbid(f, turn) for f in answer.forbid]
    passed = all(d["passed"] for d in details)
    if passed:
        headline = "; ".join(d["expected"] for d in details)
    else:
        headline = next(d["reason"] for d in details if not d["passed"])
    return EmailResult(passed=passed, max=1, headline=headline, details=details)
