
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


def _parse_sub_criteria(criteria: str) -> list[str]:
    """Split a criteria string on commas that separate independent checks.

    E.g. 'TC-{date}, TC-{date}, TC-{date}' → three TC sub-criteria.
    Commas inside braces are NOT split points.
    '&&' is also treated as a separator.
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
    return [p for p in parts if p]


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


def _event_matches_date(e: EventResponse, date_str: str) -> bool:
    """Check if an event's start date/time matches a resolved date string."""
    for fmt in ("%Y-%m-%d %I:%M %p", "%Y-%m-%d %I%p", "%Y-%m-%d %H:%M",
                "%Y-%m-%d", "%m/%d/%Y", "%B %d, %Y"):
        try:
            target = datetime.strptime(date_str.strip(), fmt)
            if fmt in ("%Y-%m-%d", "%m/%d/%Y", "%B %d, %Y"):
                return e.start.date() == target.date()
            return (e.start.date() == target.date()
                    and e.start.hour == target.hour
                    and e.start.minute == target.minute)
        except ValueError:
            continue
    return True


def _check_criteria(criteria: str, calendar: CalendarResponse, todo: list[TodoResponse]) -> bool:
    """Check a criteria string against tool state.

    Supports:
    - Prefix checks: TC (todo created), CC (calendar event created),
      RS (reschedule — event exists).
    - Content matching: TC-item3 requires a todo whose title/description
      contains 'item3'.
    - Date matching: CC-{resolved-date} requires an event starting on that
      date (only when the token has been resolved by the engine).
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

    # --- CC checks ---
    if cc_subs:
        if len(calendar.events) < len(cc_subs):
            passed = False
        for sub in cc_subs:
            date_str = _extract_date_token(sub)
            if date_str and not any(_event_matches_date(e, date_str) for e in calendar.events):
                passed = False

    # --- RS checks ---
    if rs_subs:
        if len(calendar.events) == 0:
            passed = False

    # Criteria that don't start with any known prefix but aren't "No action"
    # fall through as passed=True (lenient on free-text criteria like
    # 'Flag, insufficient info', 'Delegate Task', etc.)

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
            max_score  — total possible points
            details    — list of per-criteria results
    """
    if isinstance(input, Email):
        criteria_list = [input.success_criteria] if input.success_criteria else []
    elif isinstance(input, Scenario):
        criteria_list = input.success_criteria
    else:
        raise TypeError(f"Expected Email or Scenario, got {type(input).__name__}")

    details = []
    for criteria in criteria_list:
        passed = _check_criteria(criteria, calendar, todo)
        details.append({
            "criteria": criteria,
            "passed": passed,
        })

    if grade_by_scenario:
        all_passed = all(d["passed"] for d in details) if details else False
        return {
            "score": 1 if all_passed else 0,
            "max_score": 1 if criteria_list else 0,
            "details": details,
        }

    score = sum(1 for d in details if d["passed"])
    return {
        "score": score,
        "max_score": len(criteria_list),
        "details": details,
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
