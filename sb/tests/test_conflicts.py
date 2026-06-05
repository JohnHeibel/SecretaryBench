"""Tests for sb.conflicts — the cross-storyline calendar time-conflict measurement.

Grading is node-scoped, so these conflicts never change a score; this just verifies the
measurement itself: it replays create/move/cancel correctly and counts only true
cross-storyline clock overlaps.
"""
from datetime import date, datetime

from sb.conflicts import find_conflicts, report, _overlap, Slot
from sb.resolver import TimeInterval
from sb.schema import build_corpus
from sb.scheduler import build_plan

START = date(2026, 6, 1)


def _ti(h1, h2):
    return TimeInterval(datetime(2026, 6, 1, h1), datetime(2026, 6, 1, h2))


def _slot(node, name, h1, h2):
    return Slot(node=node, email_id=f"{node}.{name}", name=name, span=_ti(h1, h2))


def test_cross_node_overlap_detected():
    a, b = _slot("n1", "x", 9, 10), _slot("n2", "y", 9, 11)
    assert len(find_conflicts([a, b])) == 1


def test_same_node_overlap_excluded():
    # Two overlapping events in ONE storyline are that author's own business, never a conflict.
    a, b = _slot("n1", "x", 9, 10), _slot("n1", "y", 9, 11)
    assert find_conflicts([a, b]) == []


def test_touching_edges_do_not_conflict():
    assert _overlap(_ti(9, 10), _ti(10, 11)) is False


def test_point_inside_span_conflicts():
    point = TimeInterval(datetime(2026, 6, 1, 9, 30), datetime(2026, 6, 1, 9, 30))
    assert _overlap(_ti(9, 10), point) is True


def _two_authors_same_slot():
    # Two storylines that both book the SAME clock slot on the same weekday -> a guaranteed overlap.
    mk = lambda nid: {
        "id": nid, "cast": {"CEO": "you", "X": "Someone"},
        "emails": [{
            "id": f"{nid}.e1", "from": "X", "to": "CEO", "subject": "sync",
            "body": "put the sync on next:MON at 9", "depends_on": [],
            "answer": {"ops": [{"create": "sync", "kind": "event", "on": {"eq": "next:MON @09:00-10:00"}}]},
        }],
    }
    return build_corpus([mk("alpha"), mk("beta")])


def test_report_counts_a_real_conflict():
    corpus = _two_authors_same_slot()
    plan = build_plan(corpus, start_date=START, seed=42, n_days=30)
    r = report(corpus, plan)
    assert r.timed_slots == 2
    assert len(r.conflicts) == 1
    assert r.conflicting_slots == 2
    assert r.conflict_rate == 1.0


def test_cancel_vacates_slot():
    # create then cancel the same obligation -> no surviving slot, so no conflict with a twin.
    node = {
        "id": "alpha", "cast": {"CEO": "you", "X": "Someone"},
        "emails": [
            {"id": "alpha.book", "from": "X", "to": "CEO", "subject": "book", "body": "book it",
             "depends_on": [],
             "answer": {"ops": [{"create": "sync", "kind": "event", "on": {"eq": "next:MON @09:00-10:00"}}]}},
            {"id": "alpha.drop", "from": "X", "to": "CEO", "subject": "cancel", "body": "cancel it",
             "depends_on": [{"email": "alpha.book", "type": "static"}],
             "answer": {"ops": [{"cancel": "sync"}]}},
        ],
    }
    twin = {
        "id": "beta", "cast": {"CEO": "you", "X": "Someone"},
        "emails": [{"id": "beta.book", "from": "X", "to": "CEO", "subject": "book", "body": "book it",
                    "depends_on": [],
                    "answer": {"ops": [{"create": "sync", "kind": "event", "on": {"eq": "next:MON @09:00-10:00"}}]}}],
    }
    corpus = build_corpus([node, twin])
    plan = build_plan(corpus, start_date=START, seed=42, n_days=30)
    r = report(corpus, plan)
    # alpha's slot was cancelled, so only beta's survives -> no overlap.
    assert r.timed_slots == 1
    assert r.conflicts == []
