"""G5: the date-token resolver learns the forms the answer key already uses but
that it previously left unresolved (so they graded existence-only):

- weekday-of-next-week:  {nextweek-wednesday}, {nextweek-friday},
                         {nextweek-Thursday, 3pm GMT}
- weekday (this run):    {date-Tuesday}
- date+offset+time:      {date+1- 3PM}, {date+2- 11AM}, {date- this week 11pm}
- Nth-weekday-of-month:  {third Friday}, {third Friday -1 dinner},
                         {third Friday nextmonth}, {3rd Friday of November}

Pure unit tests — the resolver is a string transform, no server needed.

REF below is Wednesday, March 15, 2000, so the expected dates are hand-checkable:
next week's Monday is March 20, 2000."""
from __future__ import annotations

from datetime import datetime, timezone

import pytest

from engine import _resolve_one_token, resolve_tokens
from grader import define_grading_system
from loader import Email
from app.models.calendar import CalendarResponse, EventResponse
from app.models.todo import TodoResponse


REF = datetime(2000, 3, 15, tzinfo=timezone.utc)  # a Wednesday
assert REF.strftime("%A") == "Wednesday"


def r(tok: str, ref: datetime = REF):
    return _resolve_one_token(tok, ref)


# --- weekday of next week --------------------------------------------------

@pytest.mark.parametrize("tok,expected", [
    ("nextweek-wednesday", "March 22, 2000"),          # next week's Wed
    ("nextweek-friday", "March 24, 2000"),             # next week's Fri
    ("nextweek-Thursday", "March 23, 2000"),
    ("nextweek - wednesday", "March 22, 2000"),        # whitespace variations
    ("nextweek-Thursday, 3pm GMT", "March 23, 2000 at 03:00 PM"),  # +time, GMT dropped
    ("nextweek-thursday, 3:30pm", "March 23, 2000 at 03:30 PM"),
])
def test_nextweek_weekday(tok, expected):
    assert r(tok) == expected


def test_nextweek_weekday_from_sunday():
    # From Sunday Mar 19, "next week's Monday" is the very next day, Mar 20.
    sun = datetime(2000, 3, 19, tzinfo=timezone.utc)
    assert sun.strftime("%A") == "Sunday"
    assert r("nextweek-monday", sun) == "March 20, 2000"


def test_nextweek_weekday_unparseable_trailing_is_unresolved():
    # weekday recognized, but trailing text isn't a time -> leave unresolved
    # (grader falls back to existence-only rather than guessing).
    assert r("nextweek-friday foobar") is None


# --- weekday this run ({date-<weekday>}) -----------------------------------

def test_date_weekday_next_future_occurrence():
    assert r("date-Tuesday") == "March 21, 2000"   # next Tue after Wed 15th
    assert r("date-Wednesday") == "March 15, 2000"  # today counts (on/after)
    assert r("date-Sunday") == "March 19, 2000"


# --- date + offset + time --------------------------------------------------

@pytest.mark.parametrize("tok,expected", [
    ("date+1- 3PM", "March 16, 2000 at 03:00 PM"),
    ("date+2- 11AM", "March 17, 2000 at 11:00 AM"),
    ("date+1-3pm", "March 16, 2000 at 03:00 PM"),     # already-tight spacing
    ("date- this week 11pm", "March 15, 2000 at 11:00 PM"),
    ("date-this-week 9:45am", "March 15, 2000 at 09:45 AM"),
])
def test_date_offset_time(tok, expected):
    assert r(tok) == expected


def test_bare_thisweek_still_today():
    # adding the time form must not break the no-time case
    assert r("date-thisweek") == "March 15, 2000"
    assert r("date- this week") == "March 15, 2000"


def test_date_plus_offset_time_garbage_falls_through():
    # not a time after the dash -> not a date+time combo -> unresolved
    assert r("date+1-banana") is None


# --- Nth weekday of month --------------------------------------------------

@pytest.mark.parametrize("tok,expected", [
    ("third Friday", "March 17, 2000"),               # this month, on/after 15th
    ("third Friday -1 dinner", "March 16, 2000"),     # offset + trailing label
    ("third Friday nextmonth", "April 21, 2000"),
    ("third Friday nextmonth -1 dinner", "April 20, 2000"),
    ("3rd Friday of November", "November 17, 2000"),
    ("first Monday", "April 03, 2000"),               # March's 1st Mon (6th) is past
    ("last Friday", "March 31, 2000"),
    ("2nd Tuesday", "April 11, 2000"),                # March's 2nd Tue (14th) is past
])
def test_nth_weekday(tok, expected):
    assert r(tok) == expected


def test_nth_weekday_rolls_past_months_without_fifth():
    # From April 1, 2000: April has only 4 Fridays, May has only 4 -> June 30.
    apr = datetime(2000, 4, 1, tzinfo=timezone.utc)
    assert r("fifth Friday", apr) == "June 30, 2000"


def test_nth_weekday_bad_month_is_unresolved():
    assert r("third Friday of smarch") is None


# --- end-to-end: resolved tokens now drive strict grading ------------------

def _cal_event(dt):
    return CalendarResponse(calendar_id="c", start_date=REF, events=[
        EventResponse(event_id="e", title="m", start=dt,
                      end=dt.replace(hour=(dt.hour + 1) % 24), scenario_id=1)])


def _grade_cc(criteria, cal):
    return define_grading_system(
        Email(email_number=1, subject="s", body="b", sender="x", recipients=["y"],
              success_criteria=resolve_tokens(criteria, REF, keep_braces=True)),
        cal, [], grade_by_scenario=True)["score"]


def test_nextweek_weekday_grades_strictly():
    right = _cal_event(datetime(2000, 3, 22, 9, 0, tzinfo=timezone.utc))   # Wed next week
    wrong = _cal_event(datetime(2000, 3, 15, 9, 0, tzinfo=timezone.utc))   # served day
    assert _grade_cc("CC-{nextweek-wednesday}", right) == 1
    assert _grade_cc("CC-{nextweek-wednesday}", wrong) == 0


def test_nextweek_weekday_with_time_checks_time():
    right = _cal_event(datetime(2000, 3, 23, 15, 0, tzinfo=timezone.utc))  # Thu 3pm
    wrong_time = _cal_event(datetime(2000, 3, 23, 9, 0, tzinfo=timezone.utc))
    assert _grade_cc("CC-{nextweek-Thursday, 3pm GMT}", right) == 1
    assert _grade_cc("CC-{nextweek-Thursday, 3pm GMT}", wrong_time) == 0


def test_nth_weekday_grades_strictly():
    right = _cal_event(datetime(2000, 3, 17, 9, 0, tzinfo=timezone.utc))
    wrong = _cal_event(datetime(2000, 3, 10, 9, 0, tzinfo=timezone.utc))
    assert _grade_cc("CC-{third Friday}", right) == 1
    assert _grade_cc("CC-{third Friday}", wrong) == 0
