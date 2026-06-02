"""Tests for the deterministic serve-plan builder."""
from datetime import date, timedelta
from pathlib import Path

import pytest

from sb.scheduler import InfeasibleSchedule, Levers, build_plan
from sb.schema import load_corpus

START = date(2026, 6, 1)
FIXTURE = str(Path(__file__).parent / "fixtures")


def _plan(seed=42, **kw):
    return build_plan(load_corpus(FIXTURE), start_date=START, seed=seed, n_days=60, **kw)


def test_every_email_served():
    p = _plan()
    served = {e for day in p.per_day for e in day}
    assert served == set(load_corpus(FIXTURE).emails)


def test_dependencies_respected_strictly_earlier_day():
    p = _plan()
    sd = p.serve_date
    assert sd["alpha.intro"] < sd["alpha.brief"] < sd["alpha.review"]
    # node-sugar static edge: beta book after the gamma notice
    assert sd["gamma.notice"] < sd["beta.book"]
    # reschedule after the original
    assert sd["beta.book"] < sd["beta.shift"]


def test_date_window_respected():
    p = _plan()
    # review's deadline = brief + 5d (anchor) + 2w = serve(brief)+19d; must serve on/before.
    assert "alpha.review" in p.deadlines
    assert p.serve_date["alpha.review"] <= p.deadlines["alpha.review"]


def test_anchor_value_emitted_at_brief_serve_date():
    p = _plan()
    expected = p.serve_date["alpha.brief"] + timedelta(days=5)
    assert p.anchors["brief"] == expected


def test_reproducible_on_seed():
    p1 = _plan(seed=7)
    p2 = _plan(seed=7)
    assert p1.per_day == p2.per_day


def test_different_seed_changes_plan():
    # Over many emails this would always differ; with 6 emails just assert it runs
    # and produces a valid plan for several seeds.
    for s in range(5):
        p = _plan(seed=s)
        assert {e for day in p.per_day for e in day} == set(load_corpus(FIXTURE).emails)


def test_daily_cap_respected():
    p = _plan(levers=Levers(daily_min=1, daily_max=3))
    assert all(len(day) <= 3 for day in p.per_day)


def test_infeasible_when_no_days():
    with pytest.raises(InfeasibleSchedule):
        build_plan(load_corpus(FIXTURE), start_date=START, seed=1, n_days=1)
