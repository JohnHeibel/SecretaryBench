"""G2: the grader's TC (todo) branch now checks the due_date when the criterion
names a concrete deadline — mirroring the CC date check. Bare/unresolved tokens
stay existence-only, so this can only *tighten* grading, never make it lenient.

Pure unit tests — the grader reads in-memory models, no server needed."""
from __future__ import annotations

from datetime import datetime, timezone

from engine import resolve_tokens
from grader import define_grading_system
from loader import Email
from app.models.calendar import CalendarResponse
from app.models.todo import TodoResponse


REF = datetime(2000, 3, 15, tzinfo=timezone.utc)
_EMPTY_CAL = CalendarResponse(calendar_id="c", start_date=REF, events=[])


def _todo(due: datetime, title: str = "task", desc: str | None = None) -> TodoResponse:
    return TodoResponse(id="t1", title=title, description=desc, due_date=due,
                        created_at=REF, scenario_id=1)


def _grade_tc(criteria: str, todos: list[TodoResponse]) -> int:
    return define_grading_system(
        Email(email_number=1, subject="s", body="b", sender="x", recipients=["y"],
              success_criteria=resolve_tokens(criteria, REF, keep_braces=True)),
        _EMPTY_CAL, todos, grade_by_scenario=True)["score"]


# --- the headline: a deadline in the criterion is now enforced -------------

def test_tc_correct_due_date_passes():
    assert _grade_tc("TC-{date+2}", [_todo(datetime(2000, 3, 17, 17, 0, tzinfo=timezone.utc))]) == 1


def test_tc_wrong_due_date_fails():
    # todo exists, but due 3 days off the {date+2} target -> now fails (used to pass)
    assert _grade_tc("TC-{date+2}", [_todo(datetime(2000, 3, 20, 17, 0, tzinfo=timezone.utc))]) == 0


def test_tc_resolver_recovered_token_enforced():
    # {nextweek-friday} -> March 24; a todo due then passes, any other day fails.
    assert _grade_tc("TC-{nextweek-friday}", [_todo(datetime(2000, 3, 24, 12, 0, tzinfo=timezone.utc))]) == 1
    assert _grade_tc("TC-{nextweek-friday}", [_todo(datetime(2000, 3, 22, 12, 0, tzinfo=timezone.utc))]) == 0


# --- leniency is preserved where there's no concrete date ------------------

def test_bare_unresolved_token_stays_existence_only():
    # An unresolved token (not run through the engine) -> any todo satisfies it.
    result = define_grading_system(
        Email(email_number=1, subject="s", body="b", sender="x", recipients=["y"],
              success_criteria="TC-{nextweek-wednesday}"),  # deliberately unresolved
        _EMPTY_CAL, [_todo(datetime(2000, 9, 9, tzinfo=timezone.utc))], grade_by_scenario=True)
    assert result["score"] == 1


def test_non_date_braced_token_stays_lenient():
    # TC-{C}: "C" is a placeholder, not a date -> existence-only.
    assert _grade_tc("TC-{C}", [_todo(datetime(2002, 2, 2, tzinfo=timezone.utc))]) == 1


def test_literal_content_token_unaffected():
    # TC-greenlight: content match, no date involved.
    assert _grade_tc("TC-greenlight", [_todo(REF, title="Greenlight product A")]) == 1
    assert _grade_tc("TC-greenlight", [_todo(REF, title="something else")]) == 0


# --- composition: count, content, and date together -----------------------

def test_missing_todo_still_fails_on_count():
    assert _grade_tc("TC-{date+2}", []) == 0


def test_or_alternatives_pass_on_any_branch():
    # "TC-{A} or TC-{B}" passes if a todo's due_date matches either branch.
    crit = "TC-{date+1} or TC-{date+2}"
    assert _grade_tc(crit, [_todo(datetime(2000, 3, 17, tzinfo=timezone.utc))]) == 1  # +2
    assert _grade_tc(crit, [_todo(datetime(2000, 3, 16, tzinfo=timezone.utc))]) == 1  # +1
    assert _grade_tc(crit, [_todo(datetime(2000, 3, 30, tzinfo=timezone.utc))]) == 0  # neither


def test_multiple_tc_subs_require_count_and_each_date():
    # Two distinct deadlines: need two todos, one matching each date.
    crit = "TC-{date+1}, TC-{date+2}"
    good = [_todo(datetime(2000, 3, 16, tzinfo=timezone.utc)),
            _todo(datetime(2000, 3, 17, tzinfo=timezone.utc))]
    assert _grade_tc(crit, good) == 1
    # both todos on the same (wrong-for-one) day -> the +2 sub has no match
    bad = [_todo(datetime(2000, 3, 16, tzinfo=timezone.utc)),
           _todo(datetime(2000, 3, 16, tzinfo=timezone.utc))]
    assert _grade_tc(crit, bad) == 0
