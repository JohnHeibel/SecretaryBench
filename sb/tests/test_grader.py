"""Focused grader/oracle edge cases."""
from datetime import date, datetime

import pytest

from sb.engine import run
from sb.grader import NodeState, Obj, TurnDelta, grade_email
from sb.oracle import _target, oracle_model
from sb.resolver import Context, Interval
from sb.scheduler import build_plan
from sb.schema import Answer, Op, build_corpus


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


# --- fixed-time (timed) obligations ----------------------------------------

def _timed_answer():
    # next THU after Tue Jun 9 -> Thu Jun 11, 2-3 PM.
    return Answer(ops=[Op(verb="create", name="board", kind="event", match=["board"],
                          on={"eq": "next:THU @14:00-15:00"})])


def _board_event(start: datetime, end: datetime) -> NodeState:
    return NodeState(events=[Obj(kind="event", title="board meeting", when=start,
                                 email_id="n.a", end=end)])


def test_fixed_time_event_matches():
    ctx = Context(serve=date(2026, 6, 9))
    state = _board_event(datetime(2026, 6, 11, 14, 0), datetime(2026, 6, 11, 15, 0))
    assert grade_email(_timed_answer(), ctx, state, TurnDelta()).passed


def test_fixed_time_wrong_start_is_wrong_time():
    ctx = Context(serve=date(2026, 6, 9))
    state = _board_event(datetime(2026, 6, 11, 15, 0), datetime(2026, 6, 11, 16, 0))
    r = grade_email(_timed_answer(), ctx, state, TurnDelta())
    assert not r.passed
    assert r.details[0]["reason"] == "at the wrong time"


def test_fixed_time_wrong_duration_is_wrong_length():
    ctx = Context(serve=date(2026, 6, 9))
    state = _board_event(datetime(2026, 6, 11, 14, 0), datetime(2026, 6, 11, 14, 30))
    r = grade_email(_timed_answer(), ctx, state, TurnDelta())
    assert not r.passed
    assert r.details[0]["reason"] == "the wrong length"


def test_fixed_time_wrong_day_still_reads_as_wrong_day():
    ctx = Context(serve=date(2026, 6, 9))
    state = _board_event(datetime(2026, 6, 12, 14, 0), datetime(2026, 6, 12, 15, 0))
    r = grade_email(_timed_answer(), ctx, state, TurnDelta())
    assert not r.passed
    assert r.details[0]["reason"] == "on the wrong day"


def test_oracle_solves_a_timed_corpus():
    """A timed obligation (body and answer rendered from ONE @anchor) is oracle-solvable:
    the reference secretary places the event at exactly the resolved interval and scores 1.0."""
    node = {
        "id": "timed",
        "cast": {},
        "emails": [{
            "id": "timed.board",
            "from": "chief@co", "to": ["ceo@co"],
            "subject": "Board meeting",
            "body": "The board meets {!board = next:THU @14:00-15:00}.",
            "answer": {"ops": [{"create": "board", "kind": "event",
                                "match": ["board"], "on": {"eq": "@board"}}]},
        }],
    }
    corpus = build_corpus([node])
    plan = build_plan(corpus, start_date=date(2026, 6, 1), seed=1, n_days=30)
    res = run(corpus, plan, oracle_model)
    assert res.total == 1
    assert res.score() == 1.0

