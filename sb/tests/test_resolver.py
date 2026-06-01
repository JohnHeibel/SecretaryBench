"""Unit tests for the date grammar evaluator (the keystone)."""
from datetime import date

import pytest

from sb.resolver import (
    Context,
    Interval,
    ResolverError,
    human,
    render_body,
    resolve,
)

# A fixed serve anchor for determinism: Tuesday, June 9, 2026.
SERVE = date(2026, 6, 9)
assert SERVE.weekday() == 1  # TUE


def ctx(anchors=None):
    return Context(serve=SERVE, anchors=anchors or {})


# --- offsets ---------------------------------------------------------------

def test_calendar_day_offset():
    assert resolve("serve+9d", ctx()) == date(2026, 6, 18)
    assert resolve("serve-2d", ctx()) == date(2026, 6, 7)


def test_week_offset():
    assert resolve("serve+2w", ctx()) == date(2026, 6, 23)


def test_month_offset_clamps():
    # Jan 31 + 1 month clamps to Feb 28 (2026 not a leap year).
    c = Context(serve=date(2026, 1, 31))
    assert resolve("serve+1m", c) == date(2026, 2, 28)


def test_business_day_offset_skips_weekend():
    # Tue Jun 9 + 3 business days = Fri Jun 12.
    assert resolve("serve+3bd", ctx()) == date(2026, 6, 12)
    # Fri + 1 business day = Mon.
    c = Context(serve=date(2026, 6, 12))
    assert resolve("serve+1bd", c) == date(2026, 6, 15)


# --- weekday selectors -----------------------------------------------------

def test_next_weekday_is_strictly_after():
    # next Thursday after Tue Jun 9 -> Thu Jun 11.
    assert resolve("next:THU", ctx()) == date(2026, 6, 11)
    # next Tuesday is a full week out (strictly after), not today.
    assert resolve("next:TUE", ctx()) == date(2026, 6, 16)


def test_this_weekday_in_current_week():
    # Friday of the week containing Tue Jun 9 -> Fri Jun 12.
    assert resolve("this:FRI", ctx()) == date(2026, 6, 12)
    # Monday of this week is before serve -> Jun 8.
    assert resolve("this:MON", ctx()) == date(2026, 6, 8)


def test_nth_weekday_of_month():
    # The C19 gala: 3rd Friday of next month (July 2026) -> Jul 17.
    assert resolve("nth:3,FRI,+1m", ctx()) == date(2026, 7, 17)


def test_last_weekday_of_month():
    # last Friday of this month (June 2026) -> Jun 26.
    assert resolve("nth:last,FRI,0m", ctx()) == date(2026, 6, 26)


def test_day_of_month():
    assert resolve("dom:25,0m", ctx()) == date(2026, 6, 25)


# --- anchors (cross-node) --------------------------------------------------

def test_anchor_reference_and_arithmetic():
    anchors = {"signing": date(2026, 6, 14)}
    assert resolve("@signing+2w", Context(SERVE, anchors)) == date(2026, 6, 28)


def test_unknown_anchor_raises():
    with pytest.raises(ResolverError):
        resolve("@nope+1d", ctx())


# --- intervals -------------------------------------------------------------

def test_month_interval():
    v = resolve("month:0m", ctx())
    assert v == Interval(date(2026, 6, 1), date(2026, 6, 30))


def test_week_of_nested_expr():
    # week_of(serve+9d): Jun 18 is a Thursday -> Mon Jun 15 .. Sun Jun 21.
    v = resolve("week_of:(serve+9d)", ctx())
    assert v == Interval(date(2026, 6, 15), date(2026, 6, 21))
    assert v.contains(date(2026, 6, 18))
    assert not v.contains(date(2026, 6, 22))


# --- body rendering + emissions --------------------------------------------

def test_render_plain_token():
    r = render_body("Let's meet {next:THU}.", ctx())
    assert "Thursday, June 11th, 2026" in r.text
    assert r.emissions == {}


def test_render_emission_records_and_renders():
    r = render_body("Signing is locked for {!signing = serve+5d}.", ctx())
    assert r.emissions == {"signing": date(2026, 6, 14)}
    assert "Sunday, June 14th, 2026" in r.text


def test_later_token_sees_earlier_emission_in_same_body():
    body = "Signing {!signing = serve+5d}; kickoff two weeks later on {@signing+2w}."
    r = render_body(body, ctx())
    assert r.emissions["signing"] == date(2026, 6, 14)
    assert "June 28th, 2026" in r.text


def test_human_interval():
    assert human(Interval(date(2026, 6, 15), date(2026, 6, 21))) == "the week of June 15th, 2026"


# --- anchor-relative weekday selectors (from <expr>) -----------------------

def test_next_wd_from_anchor_midweek():
    # ref = Wed Jun 10, 2026; next FRI strictly after -> same-week Fri Jun 12.
    anchors = {"migration": date(2026, 6, 10)}
    assert date(2026, 6, 10).weekday() == 2  # WED
    assert resolve("next:FRI from @migration", Context(SERVE, anchors)) == date(2026, 6, 12)


def test_next_wd_from_anchor_on_that_weekday_jumps_a_week():
    # ref = Fri Jun 12; next FRI strictly after -> following Fri Jun 19.
    anchors = {"migration": date(2026, 6, 12)}
    assert date(2026, 6, 12).weekday() == 4  # FRI
    assert resolve("next:FRI from @migration", Context(SERVE, anchors)) == date(2026, 6, 19)


def test_this_wd_from_anchor_midweek():
    # ref = Wed Jun 10; this FRI = Friday of that week -> Jun 12.
    anchors = {"migration": date(2026, 6, 10)}
    assert resolve("this:FRI from @migration", Context(SERVE, anchors)) == date(2026, 6, 12)


def test_this_wd_from_anchor_on_that_weekday_is_same_day():
    # ref = Fri Jun 12; this FRI = same Friday Jun 12.
    anchors = {"migration": date(2026, 6, 12)}
    assert resolve("this:FRI from @migration", Context(SERVE, anchors)) == date(2026, 6, 12)


def test_from_expr_supports_offsets():
    # @migration is Wed Jun 10; +1w -> Wed Jun 17; next MON strictly after -> Jun 22.
    anchors = {"migration": date(2026, 6, 10)}
    assert resolve("next:MON from @migration+1w", Context(SERVE, anchors)) == date(2026, 6, 22)


def test_from_unknown_anchor_raises():
    with pytest.raises(ResolverError):
        resolve("next:FRI from @nope", ctx())


def test_from_does_not_change_serve_relative_forms():
    # No from clause: identical to current serve-relative behavior.
    assert resolve("next:THU", ctx()) == date(2026, 6, 11)
    assert resolve("this:FRI", ctx()) == date(2026, 6, 12)


# --- error handling --------------------------------------------------------

def test_trailing_junk_raises():
    with pytest.raises(ResolverError):
        resolve("serve+9d garbage", ctx())
