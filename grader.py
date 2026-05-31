
from __future__ import annotations

import re
from datetime import datetime
from typing import Union

from loader import Email, Scenario
from app.models.calendar import CalendarResponse, EventResponse
from app.models.todo import TodoResponse


_NO_ACTION = re.compile(r"no\s*action", re.IGNORECASE)
_PREFIX_RE = re.compile(r"^(TC|CC|RS)-?\s*", re.IGNORECASE)
_BRACED_TOKEN = re.compile(r"\{([^}]+)\}")
_EXPLICIT_COUNT_RE = re.compile(r"\(\s*single\b.*only\b", re.IGNORECASE)
# A leading "Label: " (e.g. "Update: ") that sits in front of a real prefix.
_LABEL_RE = re.compile(r"^\s*[A-Za-z]+:\s+(?=(?:TC|CC|RS)\b)", re.IGNORECASE)
# An "or ..." fragment: an alternative continuation of the previous sub.
_OR_CONT_RE = re.compile(r"^or\b\s*(.*)$", re.IGNORECASE)


def _parse_sub_criteria(criteria: str) -> list[str]:
    """Split a criteria string on commas that separate independent checks.

    E.g. 'TC-{date}, TC-{date}, TC-{date}' → three TC sub-criteria.
    Commas inside braces are NOT split points.
    '&&' is also treated as a separator.

    FIX-5 — repairs three ways the naive comma-split mangled REAL prefixed
    criteria into auto-passing junk (see SPRINT5_REMEDIATION Appendix C):
      - a lone '^' (a dataset placeholder) is dropped, not kept as a check;
      - an "or ..." fragment ("CC-{A}, or CC-{B}") is merged back onto the
        previous sub as an alternative, so its CC/TC prefix isn't lost and it
        isn't miscounted as a separate required check;
      - a leading label ("Update:  TC-{...}") is stripped so the prefix is at
        position 0 where _PREFIX_RE can see it.
    """
    criteria = criteria.replace("&&", ",")
    parts: list[str] = []
    depth = 0
    current: list[str] = []
    for ch in criteria:
        if ch in ("{", "("):
            depth += 1
        elif ch in ("}", ")"):
            depth = max(depth - 1, 0)
        if ch == "," and depth == 0:
            parts.append("".join(current).strip())
            current = []
        else:
            current.append(ch)
    tail = "".join(current).strip()
    if tail:
        parts.append(tail)

    cleaned: list[str] = []
    for p in parts:
        p = p.strip().strip("^").strip()  # drop stray caret markers
        if not p:
            continue
        p = _LABEL_RE.sub("", p)  # strip a leading "Update:"-style label
        m = _OR_CONT_RE.match(p)
        if m and cleaned:
            cleaned[-1] = f"{cleaned[-1]} or {m.group(1)}".strip()
            continue
        cleaned.append(p)
    return cleaned


def _is_gradeable(criteria: str) -> bool:
    """True if a criterion has something the grader can actually check —
    a TC/CC/RS prefix or a "No action" rule. Everything else is free-text
    (e.g. 'Add task', 'Delegate Task', 'Flag') and is reported as ungraded
    rather than silently auto-passing and inflating max_score (FIX-5, D2)."""
    if not criteria or not criteria.strip():
        return False
    if _NO_ACTION.search(criteria):
        return True
    return any(_PREFIX_RE.match(s) for s in _parse_sub_criteria(criteria))


def _extract_content_token(sub: str) -> str | None:
    """Return the literal content token from a TC/CC sub-criterion, or None.

    Literal means no braces — e.g. 'TC-item3' → 'item3'.
    Braced tokens like 'TC-{date}' are date placeholders, not content.
    """
    m = _PREFIX_RE.match(sub)
    if not m:
        return None
    remainder = sub[m.end():].strip()
    remainder = re.sub(r"\(.*\)", "", remainder).strip()
    if not remainder or _BRACED_TOKEN.search(remainder):
        return None
    if remainder.upper() in ("EOD", "EOW", "ASAP"):
        return None
    return remainder.lower()


def _extract_date_token(sub: str) -> str | None:
    """Return the resolved date string from a braced token, or None.

    If the token still contains unresolved placeholders (e.g. '{date}',
    '{nextweek-wednesday}') we return None — date matching only works
    after the engine resolves criteria tokens.
    """
    m = _BRACED_TOKEN.search(sub)
    if not m:
        return None
    inner = m.group(1).strip()
    if re.search(r"(date|nextweek|next\s*week|friday|monday|tuesday|wednesday|thursday|saturday|sunday|eod|EOD)", inner, re.IGNORECASE):
        return None
    return inner


def _todo_matches_content(t: TodoResponse, token: str) -> bool:
    haystack = t.title.lower()
    if t.description:
        haystack += " " + t.description.lower()
    return token in haystack


def _cc_sub_passes(sub: str, events: list[EventResponse]) -> bool:
    """Whether one CC sub-criterion is satisfied, honoring "or" alternatives.

    A sub may be a single check ("CC-{date}") or an alternative group merged by
    the splitter ("CC-{A} or CC-{B} or CC-{C}"), which passes if any branch's
    resolved date matches an event. Only branches that yield a *parseable* date
    drive a date check; a braced non-date token (e.g. the placeholder "{C}") or
    a still-unresolved token means "no date" → fall back to the count check in
    _check_criteria (an event merely has to exist). This keeps non-date tokens
    lenient while the fail-closed _event_matches_date makes real dates strict.
    """
    branches = re.split(r"\bor\b", sub, flags=re.IGNORECASE)
    dates = [d for d in (_extract_date_token(b) for b in branches)
             if d and _is_parseable_date(d)]
    if not dates:
        return True
    return any(_event_matches_date(e, d) for e in events for d in dates)


def _tc_date_sub_passes(sub: str, todos: list[TodoResponse]) -> bool:
    """Whether one TC sub-criterion's *date* is satisfied by some todo's due_date.

    Mirrors _cc_sub_passes (G2): a braced token that resolves to a real date is
    strict-matched against todo due_dates, honoring "or" alternatives; a non-date
    token (e.g. the placeholder "{C}") or a still-unresolved token (e.g. a bare
    "{date}" or a form the resolver can't read) means "no specific deadline" ->
    True, leaving existence/count and content checks to the caller. This is the
    deadline half of a TC check; content matching stays separate."""
    branches = re.split(r"\bor\b", sub, flags=re.IGNORECASE)
    dates = [d for d in (_extract_date_token(b) for b in branches)
             if d and _is_parseable_date(d)]
    if not dates:
        return True
    return any(_todo_matches_date(t, d) for t in todos for d in dates)


# Formats that carry only a date (compare date) vs a date+time (compare both).
_DATE_ONLY_FMTS = ("%Y-%m-%d", "%m/%d/%Y", "%B %d, %Y")
_DATETIME_FMTS = ("%Y-%m-%d %I:%M %p", "%Y-%m-%d %I%p", "%Y-%m-%d %H:%M",
                  # engine._DATETIME_FMT — time-of-day tokens like {date-3PM}
                  # resolve to "March 15, 2000 at 03:00 PM" (FIX-2 follow-up).
                  "%B %d, %Y at %I:%M %p")


def _is_parseable_date(s: str) -> bool:
    """True if `s` is a date/datetime in one of the formats the engine emits."""
    s = s.strip()
    for fmt in _DATE_ONLY_FMTS + _DATETIME_FMTS:
        try:
            datetime.strptime(s, fmt)
            return True
        except ValueError:
            continue
    return False


def _match_datetime(actual: datetime, date_str: str) -> bool:
    """Whether `actual` matches the resolved `date_str`.

    Date-only target formats compare the calendar date; date+time formats
    compare date, hour, and minute. Returns False when the target can't be
    parsed: by the time we get here the criterion named a specific date, so an
    unparseable / unmatched target is a non-match, not a free pass. (Previously
    this returned True, which let a wrong-date event satisfy a time-of-day
    criterion such as CC-{date-3PM} — caught by the Sprint-5 adversarial gate.)
    """
    target = date_str.strip()
    for fmt in _DATE_ONLY_FMTS:
        try:
            return actual.date() == datetime.strptime(target, fmt).date()
        except ValueError:
            continue
    for fmt in _DATETIME_FMTS:
        try:
            t = datetime.strptime(target, fmt)
            return (actual.date() == t.date()
                    and actual.hour == t.hour
                    and actual.minute == t.minute)
        except ValueError:
            continue
    return False


def _event_matches_date(e: EventResponse, date_str: str) -> bool:
    """Whether an event's start date/time matches a resolved date string."""
    return _match_datetime(e.start, date_str)


def _todo_matches_date(t: TodoResponse, date_str: str) -> bool:
    """Whether a todo's due_date matches a resolved date string (G2 — mirrors
    _event_matches_date so a deadline in the criterion is actually enforced)."""
    return _match_datetime(t.due_date, date_str)


def _check_criteria(criteria: str, calendar: CalendarResponse, todo: list[TodoResponse]) -> bool:
    """Check a criteria string against tool state.

    Supports:
    - Prefix checks: TC (todo created), CC (calendar event created),
      RS (reschedule — event exists).
    - Content matching: TC-item3 requires a todo whose title/description
      contains 'item3'.
    - Date matching: CC-{resolved-date} requires an event starting on that
      date, and TC-{resolved-date} requires a todo whose due_date is that date
      (only when the token has been resolved by the engine; bare/unresolved
      tokens stay existence-only).
    - Count awareness: 'TC-{date}, TC-{date}, TC-{date}' requires exactly
      3 todos. Parenthetical '(single to-do only, not three)' forces
      count == 1.
    - No action: nothing should have been created.
    """
    if not criteria or not criteria.strip():
        return False

    criteria = criteria.strip()

    if _NO_ACTION.search(criteria):
        return len(todo) == 0 and len(calendar.events) == 0

    subs = _parse_sub_criteria(criteria)
    tc_subs = [s for s in subs if re.match(r"TC", s, re.IGNORECASE)]
    cc_subs = [s for s in subs if re.match(r"CC", s, re.IGNORECASE)]
    rs_subs = [s for s in subs if re.match(r"RS", s, re.IGNORECASE)]

    explicit_single = bool(_EXPLICIT_COUNT_RE.search(criteria))

    passed = True

    # --- TC checks ---
    if tc_subs:
        required_count = 1 if explicit_single else len(tc_subs)
        if len(todo) < required_count:
            passed = False
        if explicit_single and len(todo) > 1:
            passed = False
        for sub in tc_subs:
            token = _extract_content_token(sub)
            if token and not any(_todo_matches_content(t, token) for t in todo):
                passed = False
            # G2: if the sub names a concrete deadline, a todo's due_date must
            # match it (mirrors the CC date check). Bare/unresolved -> lenient.
            if not _tc_date_sub_passes(sub, todo):
                passed = False

    # --- CC checks ---
    if cc_subs:
        if len(calendar.events) < len(cc_subs):
            passed = False
        for sub in cc_subs:
            if not _cc_sub_passes(sub, calendar.events):
                passed = False

    # --- RS checks ---
    if rs_subs:
        if len(calendar.events) == 0:
            passed = False

    # Sub-criteria with no known prefix (free-text like 'Flag') add no check —
    # they neither pass nor fail this criterion. A criterion that is ENTIRELY
    # free-text never reaches here: define_grading_system routes it to the
    # `ungraded` list instead of grading it (FIX-5).

    return passed


def define_grading_system(
    input: Union[Email, Scenario],
    calendar: CalendarResponse,
    todo: list[TodoResponse],
    grade_by_scenario: bool = True,
) -> dict:
    """
    Grade a single email or full scenario against the current tool state.

    Args:
        input:    Email object (single email) or Scenario object (email chain)
        calendar: CalendarResponse from Week 2 API — contains events list
        todo:     list of TodoResponse from Week 2 API — contains todo items
        grade_by_scenario: When True, all criteria must pass for 1 point (else 0).
                           When False, each criterion is scored individually.

    Returns:
        dict with keys:
            score      — points awarded
            max_score  — total possible points (gradeable criteria only)
            details    — list of per-criteria results (gradeable only)
            ungraded   — free-text criteria excluded from scoring (FIX-5)

    Free-text criteria (no TC/CC/RS prefix, not "No action") are NOT graded and
    do NOT count toward max_score; they are reported in `ungraded` so the score
    reflects only what the grader can actually verify (FIX-5, D2). See
    docs/GRADER.md.
    """
    if isinstance(input, Email):
        criteria_list = [input.success_criteria] if input.success_criteria else []
    elif isinstance(input, Scenario):
        criteria_list = input.success_criteria
    else:
        raise TypeError(f"Expected Email or Scenario, got {type(input).__name__}")

    details = []
    ungraded = []
    for criteria in criteria_list:
        if not _is_gradeable(criteria):
            ungraded.append(criteria)
            continue
        details.append({
            "criteria": criteria,
            "passed": _check_criteria(criteria, calendar, todo),
        })

    if grade_by_scenario:
        all_passed = all(d["passed"] for d in details) if details else False
        return {
            "score": 1 if all_passed else 0,
            "max_score": 1 if details else 0,
            "details": details,
            "ungraded": ungraded,
        }

    score = sum(1 for d in details if d["passed"])
    return {
        "score": score,
        "max_score": len(details),
        "details": details,
        "ungraded": ungraded,
    }


# ---------------------------------------------------------------------------
# CLI quick test
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    from datetime import datetime, timezone
    from loader import load_scenarios

    scenarios = load_scenarios("Emails.xlsx")
    print(f"Loaded {len(scenarios)} scenarios\n")

    # Create empty calendar and todo list for demo
    empty_calendar = CalendarResponse(
        calendar_id="demo",
        start_date=datetime.now(timezone.utc),
        events=[],
    )
    empty_todos: list[TodoResponse] = []

    total_score = 0
    total_max = 0

    # Grade each scenario
    for s in scenarios:
        result = define_grading_system(s, calendar=empty_calendar, todo=empty_todos)
        total_score += result["score"]
        total_max += result["max_score"]

        if result["max_score"] > 0:
            print(f"[{s.scenario_type}] {s.scenario_id}: {result['score']}/{result['max_score']}")

    print(f"\nTotal: {total_score}/{total_max}")

    # Also demo single-email grading
    print("\n--- Single email grading demo ---")
    for s in scenarios[:3]:
        for e in s.emails:
            result = define_grading_system(e, calendar=empty_calendar, todo=empty_todos)
            print(f"  Email #{e.email_number} ({e.subject[:40]}): {result['score']}/{result['max_score']}")
