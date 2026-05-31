"""FIX-5: splitter repairs + free-text "ungraded" marker.

Pure unit tests on grader internals — no server required."""
from __future__ import annotations

from datetime import datetime, timezone

from grader import _parse_sub_criteria, _is_gradeable, define_grading_system
from loader import Email
from app.models.calendar import CalendarResponse, EventResponse


def _email(criteria: str) -> Email:
    return Email(email_number=1, subject="s", body="b", sender="x",
                 recipients=["y"], success_criteria=criteria)


def _empty_cal() -> CalendarResponse:
    return CalendarResponse(calendar_id="c", start_date=datetime(2000, 1, 1, tzinfo=timezone.utc), events=[])


# --- splitter repairs ------------------------------------------------------

def test_lone_caret_is_dropped():
    assert _parse_sub_criteria("^") == []
    assert _is_gradeable("^") is False


def test_or_alternatives_merge_into_one_cc_sub():
    parsed = _parse_sub_criteria("CC-{date- 12pm}, or  CC-{date+1- 3PM}, or CC-{date+2- 11AM}")
    assert len(parsed) == 1, parsed
    assert parsed[0].startswith("CC-")
    assert parsed[0].count("CC-") == 3  # all three alternatives preserved
    assert _is_gradeable("CC-{date- 12pm}, or  CC-{date+1- 3PM}, or CC-{date+2- 11AM}")


def test_leading_label_stripped_so_prefix_is_recognized():
    assert _parse_sub_criteria("Update:  TC-{nextweek-wednesday}") == ["TC-{nextweek-wednesday}"]
    assert _is_gradeable("Update:  TC-{nextweek-wednesday}")


def test_genuine_free_text_not_gradeable():
    for s in ["Add task", "Delegate Task", "Flag", "Create to:do",
              "Find when quarterly earnings call is", "Must not re:add"]:
        assert _is_gradeable(s) is False, s


def test_commas_inside_resolved_date_braces_not_split():
    # FIX-2 produces "CC-{March 15, 2000}" — the comma must not split it.
    assert _parse_sub_criteria("CC-{March 15, 2000}") == ["CC-{March 15, 2000}"]


# --- ungraded marker -------------------------------------------------------

def test_free_text_excluded_from_max_score_and_reported():
    result = define_grading_system(_email("Add task"), _empty_cal(), [], grade_by_scenario=False)
    assert result["max_score"] == 0, result   # not counted
    assert result["score"] == 0
    assert result["ungraded"] == ["Add task"]
    assert result["details"] == []


def test_free_text_does_not_inflate_scenario_max():
    # Per-scenario: a scenario whose only criterion is free-text scores 0/0.
    from loader import Scenario
    s = Scenario(scenario_id="F", scenario_type="F", emails=[],
                 success_criteria=["Delegate Task"], puzzle_summary=None)
    result = define_grading_system(s, _empty_cal(), [], grade_by_scenario=True)
    assert result["max_score"] == 0, result
    assert result["ungraded"] == ["Delegate Task"]


def test_mixed_criterion_graded_on_real_subcheck_only():
    # "TC-item, Flag": gradeable (TC present); Flag adds no pass/fail.
    todo_match = define_grading_system(
        _email("TC-item, Flag"), _empty_cal(), [], grade_by_scenario=False)
    assert todo_match["max_score"] == 1  # counted (it has a real check)
    assert todo_match["ungraded"] == []
    assert todo_match["score"] == 0      # no todo created -> fails the TC check


def test_no_action_still_gradeable():
    result = define_grading_system(_email("No action required"), _empty_cal(), [],
                                   grade_by_scenario=False)
    assert result["max_score"] == 1
    assert result["score"] == 1  # nothing created -> No action passes
    assert result["ungraded"] == []
