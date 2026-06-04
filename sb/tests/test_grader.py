"""Focused grader/oracle edge cases."""
from datetime import date, datetime

import pytest

from sb.grader import NodeState, Obj, TurnDelta, grade_email
from sb.oracle import _target
from sb.resolver import Context, Interval
from sb.schema import Answer, Op


def test_by_predicate_rejects_due_date_before_email_arrives():
    ctx = Context(serve=date(2026, 6, 10))
    answer = Answer(ops=[Op(verb="create", name="filing", kind="todo", match=["filing"], on={"by": "serve+5d"})])
    state = NodeState(todos=[Obj(kind="todo", title="filing", when=datetime(2026, 6, 9, 17), email_id="n.a")])

    result = grade_email(answer, ctx, state, TurnDelta())

    assert not result.passed
    assert result.details[0]["reason"] == "on the wrong day"


def test_by_predicate_accepts_due_date_between_serve_and_deadline():
    ctx = Context(serve=date(2026, 6, 10))
    answer = Answer(ops=[Op(verb="create", name="filing", kind="todo", match=["filing"], on={"by": "serve+5d"})])
    state = NodeState(todos=[Obj(kind="todo", title="filing", when=datetime(2026, 6, 12, 17), email_id="n.a")])

    assert grade_email(answer, ctx, state, TurnDelta()).passed


def test_oracle_target_avoids_not_in_blackout():
    ctx = Context(serve=date(2026, 6, 1), anchors={"blackout": Interval(date(2026, 6, 8), date(2026, 6, 10))})

    assert _target({"in": "week_of:(serve+1w)", "not_in": "@blackout"}, ctx) == date(2026, 6, 11)


def test_oracle_target_errors_when_window_fully_blocked():
    ctx = Context(serve=date(2026, 6, 1), anchors={"blackout": Interval(date(2026, 6, 8), date(2026, 6, 14))})

    with pytest.raises(ValueError, match="fully blocked"):
        _target({"in": "week_of:(serve+1w)", "not_in": "@blackout"}, ctx)

