"""FIX-2: success_criteria date tokens are resolved before grading, so the
grader's date matching actually fires (correct date passes, wrong date fails).

These are pure unit tests — no server required (grader reads in-memory models,
engine resolution is a string transform)."""
from __future__ import annotations

from datetime import datetime, timezone

from engine import resolve_tokens, apply_date_substitutions
from grader import define_grading_system, _extract_date_token
from loader import Email
from app.models.calendar import CalendarResponse, EventResponse


REF = datetime(2000, 3, 15, tzinfo=timezone.utc)


def _calendar_with_event_on(day: int) -> CalendarResponse:
    start = datetime(2000, 3, day, 9, 0, tzinfo=timezone.utc)
    return CalendarResponse(
        calendar_id="c",
        start_date=REF,
        events=[EventResponse(event_id="e", title="Mtg", start=start,
                              end=start.replace(hour=10), scenario_id=1)],
    )


# --- engine resolution -----------------------------------------------------

def test_resolve_tokens_keep_braces_wraps_resolved_date():
    # criteria resolution keeps braces around the resolved date
    assert resolve_tokens("CC-{date}", REF, keep_braces=True) == "CC-{March 15, 2000}"
    # subject/body resolution strips braces (unchanged behavior)
    assert resolve_tokens("meet {date}", REF) == "meet March 15, 2000"


def test_resolve_tokens_keep_braces_leaves_unknown_tokens():
    # unknown / range tokens stay raw so the grader still skips them
    assert resolve_tokens("CC-{Annual Report 1}", REF, keep_braces=True) == "CC-{Annual Report 1}"


def test_apply_date_substitutions_resolves_criteria():
    e = Email(email_number=1, subject="s", body="b", sender="x", recipients=["y"],
              success_criteria="CC-{date}")
    resolved = apply_date_substitutions(e, REF)
    assert resolved.success_criteria == "CC-{March 15, 2000}"
    assert e.success_criteria == "CC-{date}"  # original untouched


def test_no_unresolved_date_tokens_after_resolution():
    # The criterion handed to the grader must not still be a bare {date}-style token.
    e = Email(email_number=1, subject="s", body="b", sender="x", recipients=["y"],
              success_criteria="CC-{date}, TC-{date+2}")
    resolved = apply_date_substitutions(e, REF).success_criteria
    for unresolved in ("{date}", "{date+2}"):
        assert unresolved not in resolved, resolved
    assert _extract_date_token("CC-{March 15, 2000}") == "March 15, 2000"


# --- grader discrimination (the headline acceptance) -----------------------

def test_correct_date_passes():
    cal = _calendar_with_event_on(15)
    result = define_grading_system(
        Email(email_number=1, subject="s", body="b", sender="x", recipients=["y"],
              success_criteria=resolve_tokens("CC-{date}", REF, keep_braces=True)),
        cal, [], grade_by_scenario=True,
    )
    assert result["score"] == 1, result


def test_wrong_date_fails():
    cal = _calendar_with_event_on(20)  # event on the 20th, criterion wants the 15th
    result = define_grading_system(
        Email(email_number=1, subject="s", body="b", sender="x", recipients=["y"],
              success_criteria=resolve_tokens("CC-{date}", REF, keep_braces=True)),
        cal, [], grade_by_scenario=True,
    )
    assert result["score"] == 0, result


def _cal_with_event(dt):
    return CalendarResponse(
        calendar_id="c", start_date=REF,
        events=[EventResponse(event_id="e", title="m", start=dt,
                              end=dt.replace(hour=(dt.hour + 1) % 24), scenario_id=1)])


def _grade(criteria, cal):
    return define_grading_system(
        Email(email_number=1, subject="s", body="b", sender="x", recipients=["y"],
              success_criteria=resolve_tokens(criteria, REF, keep_braces=True)),
        cal, [], grade_by_scenario=True)["score"]


# --- time-of-day tokens (adversarial-gate regression) ----------------------

def test_time_of_day_token_wrong_date_fails():
    # CC-{date-3PM} -> "March 15, 2000 at 03:00 PM"; an event in 2005 must FAIL.
    wrong = _cal_with_event(datetime(2005, 1, 1, 15, 0, tzinfo=timezone.utc))
    assert _grade("CC-{date-3PM}", wrong) == 0


def test_time_of_day_token_correct_date_and_time_passes():
    right = _cal_with_event(datetime(2000, 3, 15, 15, 0, tzinfo=timezone.utc))
    assert _grade("CC-{date-3PM}", right) == 1


def test_time_of_day_token_correct_date_wrong_time_fails():
    # right date, wrong time (9am vs 3pm) -> the time-specific criterion fails.
    wrong_time = _cal_with_event(datetime(2000, 3, 15, 9, 0, tzinfo=timezone.utc))
    assert _grade("CC-{date-3PM}", wrong_time) == 0


def test_non_date_braced_token_stays_lenient():
    # CC-{C}: "C" is a non-date placeholder -> no date check, count-only. An
    # event existing satisfies it (must NOT regress to always-fail).
    cal = _cal_with_event(datetime(2001, 6, 6, 9, 0, tzinfo=timezone.utc))
    assert _grade("CC-{C}", cal) == 1


def test_unresolved_token_still_skips_date_check():
    # If a criterion were NOT resolved, the grader must not crash; it skips the
    # date sub-check and falls back to the count check (event exists -> pass).
    cal = _calendar_with_event_on(20)
    result = define_grading_system(
        Email(email_number=1, subject="s", body="b", sender="x", recipients=["y"],
              success_criteria="CC-{date}"),  # deliberately unresolved
        cal, [], grade_by_scenario=True,
    )
    # _extract_date_token returns None for "{date}", so only the count check runs.
    assert result["score"] == 1, result
