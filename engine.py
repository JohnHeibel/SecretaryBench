from __future__ import annotations

import calendar as _calendar
import json
import os
import re
import sys
import threading
import time
from collections import defaultdict
from copy import deepcopy
from datetime import datetime, timedelta, timezone
from typing import Optional

from loader import load_scenarios, Email, Scenario
from flow_controller import FlowController
from grader import define_grading_system
from pipeline import register_scenario, fetch_scenario_results, scenario_str_to_int
from app.models.calendar import CalendarResponse
from harness.config import get_harness as _get_harness

TURN_LOG_PATH = os.environ.get("TURN_LOG_PATH", "turn_results.jsonl")
_turn_log_lock = threading.Lock()


def _log_turn(result, sim_date: datetime, email_index: int) -> None:
    if not TURN_LOG_PATH or result is None:
        return
    row = {
        "ts": datetime.now(timezone.utc).isoformat(),
        "sim_date": sim_date.strftime("%Y-%m-%d"),
        "scenario_id": result.scenario_id,
        "email_index": email_index,
        "elapsed_s": round(result.elapsed_s, 2),
        "rounds": result.rounds,
        "input_tokens": result.input_tokens,
        "output_tokens": result.output_tokens,
        "tool_calls": result.tool_calls,
        "harness": os.environ.get("HARNESS", "claude-p"),
    }
    try:
        with _turn_log_lock, open(TURN_LOG_PATH, "a") as f:
            f.write(json.dumps(row) + "\n")
    except OSError:
        pass


# Mirrors grader._NO_ACTION. Used to override the per-email grade when the AI
# took a destructive action (delete_*) that doesn't appear in the added/modified
# diff — the grader's len(items)==0 check would otherwise falsely pass.
_NO_ACTION_RE = re.compile(r"no\s*action", re.IGNORECASE)


SIM_START = datetime(2000, 1, 1, tzinfo=timezone.utc)
SIM_DAYS = 100

_TOKEN_RE = re.compile(r"\{([^}]+)\}")
_DATE_FMT = "%B %d, %Y"
_DATETIME_FMT = "%B %d, %Y at %I:%M %p"


def _add_months(d: datetime, months: int) -> datetime:
    """Add `months` to `d`, clamping the day to the last valid day of the result month."""
    month_index = d.month - 1 + months
    year = d.year + month_index // 12
    month = month_index % 12 + 1
    last_day = _calendar.monthrange(year, month)[1]
    day = min(d.day, last_day)
    return d.replace(year=year, month=month, day=day)


def _parse_time(s: str) -> Optional[tuple[int, int]]:
    """Parse '10am', '2pm', '1:14pm', '12:30pm' -> (hour, minute) in 24h, or None."""
    m = re.match(r"^(\d{1,2})(?::(\d{2}))?(am|pm)$", s)
    if not m:
        return None
    hour = int(m.group(1))
    minute = int(m.group(2) or 0)
    suffix = m.group(3)
    if suffix == "pm" and hour != 12:
        hour += 12
    elif suffix == "am" and hour == 12:
        hour = 0
    return hour, minute


def _next_day_of_month(sim_date: datetime, target_dom: int) -> datetime:
    """Return the next future occurrence of `target_dom` on or after sim_date."""
    try:
        candidate = sim_date.replace(day=target_dom)
    except ValueError:
        candidate = _add_months(sim_date, 1)
        last_day = _calendar.monthrange(candidate.year, candidate.month)[1]
        candidate = candidate.replace(day=min(target_dom, last_day))
    if candidate < sim_date:
        nxt = _add_months(sim_date, 1)
        last_day = _calendar.monthrange(nxt.year, nxt.month)[1]
        candidate = nxt.replace(day=min(target_dom, last_day))
    return candidate


def _resolve_one_token(raw: str, sim_date: datetime) -> Optional[str]:
    """Try to resolve a single token (the text inside {...}).

    Returns the resolved string, or None if unrecognized so the caller can
    leave the original token untouched.
    """
    # Normalize: lowercase + strip every whitespace char so "{date- next week}"
    # and "{date-nextweek}" collapse to the same form.
    normalized = re.sub(r"\s+", "", raw.lower())

    if normalized == "date":
        return sim_date.strftime(_DATE_FMT)

    if normalized in ("meeting-link", "link"):
        return "[meeting link]"

    # {date+N} or {date+Nweeks}
    m = re.match(r"^date\+(\d+)(weeks?)?$", normalized)
    if m:
        n = int(m.group(1))
        delta = timedelta(weeks=n) if m.group(2) else timedelta(days=n)
        return (sim_date + delta).strftime(_DATE_FMT)

    if normalized in ("date-tomorrow", "date-nextday"):
        return (sim_date + timedelta(days=1)).strftime(_DATE_FMT)

    if normalized == "date-beginningmonth":
        return sim_date.replace(day=1).strftime(_DATE_FMT)

    # {date-nextmonth} / {date-nextmonth+N}: same day-of-month next month, +N
    m = re.match(r"^date-nextmonth(?:\+(\d+))?$", normalized)
    if m:
        base = _add_months(sim_date, 1)
        if m.group(1):
            base += timedelta(days=int(m.group(1)))
        return base.strftime(_DATE_FMT)

    # All "next week" + relative offset variants:
    #   {date-nextweek}, {date-next-week}, {date- next week},
    #   {nextweek-date}, {nextweek - date}, {nextweek-date +3},
    #   {nextweek -date +5}, {nextweek - date -2}, ...
    # After whitespace strip they collapse onto one of these shapes.
    m = re.match(r"^(?:date-?next-?week|nextweek-?date)([+-]\d+)?$", normalized)
    if m:
        base = sim_date + timedelta(weeks=1)
        if m.group(1):
            base += timedelta(days=int(m.group(1)))
        return base.strftime(_DATE_FMT)

    # {date-thisweek} / {date-this-week} -> ambiguous, treat as today
    if re.match(r"^date-?this-?week$", normalized):
        return sim_date.strftime(_DATE_FMT)

    # {date-14th}, {date-25th} -> next future occurrence of that day-of-month
    m = re.match(r"^date-(\d{1,2})(?:st|nd|rd|th)$", normalized)
    if m:
        return _next_day_of_month(sim_date, int(m.group(1))).strftime(_DATE_FMT)

    # {date-10AM}, {date-1:14PM}, {date-12:30PM} -> today at that time
    m = re.match(r"^date-(.+)$", normalized)
    if m:
        time_parsed = _parse_time(m.group(1))
        if time_parsed is not None:
            hour, minute = time_parsed
            return sim_date.replace(hour=hour, minute=minute, second=0, microsecond=0)\
                .strftime(_DATETIME_FMT)

    return None


def resolve_tokens(text: str, sim_date: datetime) -> str:
    """Replace every {...} date/link token in `text` with a resolved value.

    Tokens that don't match a known pattern (e.g. `{Annual Report 1}`,
    `{contract 2}`, range tokens like `{date-12:30-2:00PM}`) are left as-is
    so they remain visible for downstream review.
    """
    if not text:
        return text

    def _sub(match: re.Match) -> str:
        resolved = _resolve_one_token(match.group(1), sim_date)
        return resolved if resolved is not None else match.group(0)

    return _TOKEN_RE.sub(_sub, text)


def apply_date_substitutions(email: Email, sim_date: datetime) -> Email:
    """Return a copy of `email` with subject and body tokens resolved."""
    e = deepcopy(email)
    e.subject = resolve_tokens(e.subject, sim_date)
    e.body = resolve_tokens(e.body, sim_date)
    return e


def model_interaction_mock(
    emails: list[Email], sim_date: datetime, verbose: bool = True
) -> None:
    """Placeholder for the real model call. Sleeps briefly to simulate latency."""
    if verbose:
        import bench_logger as log
        log.info("model", f"serving {len(emails)} email(s) on {sim_date.strftime('%Y-%m-%d')}")
    time.sleep(0.05)


def _state_diff(before: dict, after: dict) -> dict:
    """Compute added/removed/modified todos and events between two snapshots.

    Both inputs are the dict returned by fetch_scenario_results — {"calendar":
    CalendarResponse, "todos": list[TodoResponse]}. Items are identified by
    `id` (todos) and `event_id` (events). Modified = same id, different
    serialized content.

    With the current tool surface (create/update/delete all visible), all three
    operations need to be detected to attribute a single email's effect.
    """
    before_todos = {t.id: t for t in before["todos"]}
    after_todos = {t.id: t for t in after["todos"]}
    before_events = {e.event_id: e for e in before["calendar"].events}
    after_events = {e.event_id: e for e in after["calendar"].events}

    added_todos = [t for tid, t in after_todos.items() if tid not in before_todos]
    removed_todos = [t for tid, t in before_todos.items() if tid not in after_todos]
    modified_todos = [t for tid, t in after_todos.items()
                      if tid in before_todos
                      and t.model_dump() != before_todos[tid].model_dump()]

    added_events = [e for eid, e in after_events.items() if eid not in before_events]
    removed_events = [e for eid, e in before_events.items() if eid not in after_events]
    modified_events = [e for eid, e in after_events.items()
                       if eid in before_events
                       and e.model_dump() != before_events[eid].model_dump()]

    return {
        "added_todos": added_todos,
        "removed_todos": removed_todos,
        "modified_todos": modified_todos,
        "added_events": added_events,
        "removed_events": removed_events,
        "modified_events": modified_events,
    }


def _diff_took_action(diff: dict) -> bool:
    return any((
        diff["added_todos"], diff["removed_todos"], diff["modified_todos"],
        diff["added_events"], diff["removed_events"], diff["modified_events"],
    ))


def _grade_email_against_diff(
    email: Email, before: dict, after: dict
) -> dict:
    """Grade one email's success_criteria against the items the AI created or
    modified during that email's turn.

    Why diff-based: the calendar/todos store is cumulative within a scenario,
    so naive grading at email-N time would let earlier emails' actions falsely
    satisfy or falsely violate email-N's criteria. The diff isolates this turn.

    Removed items intentionally don't appear in the synthetic state passed to
    the grader (the grader's content/date checks don't have a sensible reading
    for "was deleted"). They DO show up via _diff_took_action — used below to
    override "No action" grades when the AI deleted something.
    """
    diff = _state_diff(before, after)
    synthetic_calendar = CalendarResponse(
        calendar_id="diff",
        start_date=after["calendar"].start_date,
        events=diff["added_events"] + diff["modified_events"],
    )
    synthetic_todos = diff["added_todos"] + diff["modified_todos"]

    result = define_grading_system(
        email, synthetic_calendar, synthetic_todos, grade_by_scenario=False
    )

    # If the AI took ANY action (including a delete that's invisible in the
    # synthetic state), flip any "No action" sub-criteria that the grader
    # incorrectly passed. Then recompute score.
    if _diff_took_action(diff):
        flipped = False
        for d in result["details"]:
            if d["passed"] and _NO_ACTION_RE.search(d["criteria"]):
                d["passed"] = False
                flipped = True
        if flipped:
            result["score"] = sum(1 for d in result["details"] if d["passed"])

    return result


def run_simulation(
    path: str = "Emails.xlsx",
    sim_days: int = SIM_DAYS,
    sim_start: datetime = SIM_START,
    seed: Optional[int] = 42,
    verbose: bool = True,
    model_fn=None,
    scenarios: list[Scenario] | None = None,
    grade_by_scenario: bool = True,
) -> dict:
    """Run the full N-day benchmark simulation and return aggregated results."""
    if scenarios is None:
        scenarios = load_scenarios(path)
    controller = FlowController(scenarios, seed=seed)
    controller.build_schedule(total_days=sim_days)

    sim_date = sim_start
    total_score = 0
    total_max = 0
    email_total_score = 0
    email_total_max = 0
    daily_log: list[dict] = []
    # Per-scenario-type breakdown so we can see which categories the model
    # handles well vs poorly. Same shape across model swaps, so a Haiku vs
    # Sonnet comparison is just diffing two of these.
    by_type: dict[str, dict[str, int]] = defaultdict(
        lambda: {"count": 0, "score": 0, "max_score": 0}
    )
    # Parallel breakdown for per-email grading. Same shape; "count" here is
    # the number of emails graded (vs scenarios in by_type), so it'll be
    # higher.
    by_type_email: dict[str, dict[str, int]] = defaultdict(
        lambda: {"count": 0, "score": 0, "max_score": 0}
    )


    _model_harness = _get_harness() if model_fn is None else None

    if verbose:
        import bench_logger as log
        log.sim_header(len(scenarios), sim_days, sim_start.strftime("%Y-%m-%d"))

    for day in range(sim_days):
        ready = controller.step(day)

        # Seed the FastAPI store for any scenarios that just activated. This is
        # Eyasu's adapter (pipeline.register_scenario) — it converts the loader
        # objects into the store's pydantic shapes and inserts them so the AI
        # can read the new emails through the API.
        for activated in controller.newly_activated_today():
            register_scenario(activated, sim_date)

        if ready and verbose:
            log.day_header(day + 1, sim_date.strftime("%Y-%m-%d"), len(ready))

        for email, scenario, idx in ready:
            resolved = apply_date_substitutions(email, sim_date)

            # Snapshot state BEFORE the model acts, so we can diff and attribute
            # what changed to this email specifically (not to earlier emails
            # in the chain).
            state_before = fetch_scenario_results(scenario)

            if model_fn is not None:
                model_fn(resolved, sim_date)
            elif _model_harness is not None:
                turn_result = _model_harness.run_turn(resolved, sim_date, scenario_str_to_int(scenario.scenario_id))
                _log_turn(turn_result, sim_date, idx + 1)
            else:
                model_interaction_mock([resolved], sim_date, verbose=verbose)

            # Grade THIS email against the diff before the controller advances
            # state. Skip silently when the email has no criteria — common for
            # informational chain links that don't define their own pass/fail.
            if resolved.success_criteria:
                state_after = fetch_scenario_results(scenario)
                email_result = _grade_email_against_diff(
                    resolved, state_before, state_after
                )
                email_total_score += email_result["score"]
                email_total_max += email_result["max_score"]
                stype = scenario.scenario_type or "(untyped)"
                eb = by_type_email[stype]
                eb["count"] += 1
                eb["score"] += email_result["score"]
                eb["max_score"] += email_result["max_score"]

            controller.mark_served(
                scenario.scenario_id, idx,
                day=day, sim_date=sim_date.strftime("%Y-%m-%d"),
            )

        # Grade scenarios that completed today (per-scenario, matches grader design)
        day_score = 0
        day_max = 0
        for scenario in controller.completed_scenarios_today():
            state = fetch_scenario_results(scenario)
            result = define_grading_system(scenario, state["calendar"], state["todos"],
                                          grade_by_scenario=grade_by_scenario)
            day_score += result["score"]
            day_max += result["max_score"]
            stype = scenario.scenario_type or "(untyped)"
            bucket = by_type[stype]
            bucket["count"] += 1
            bucket["score"] += result["score"]
            bucket["max_score"] += result["max_score"]
            if verbose and result["max_score"] > 0:
                log.grade_result(scenario.scenario_type, scenario.scenario_id,
                                 result["score"], result["max_score"],
                                 details=result.get("details"))
            if _model_harness is not None:
                _model_harness.scenario_completed(scenario_str_to_int(scenario.scenario_id))

        total_score += day_score
        total_max += day_max
        if ready or day_max:
            daily_log.append({
                "day": day + 1,
                "date": sim_date.strftime("%Y-%m-%d"),
                "served": len(ready),
                "score": day_score,
                "max_score": day_max,
            })

        sim_date += timedelta(days=1)

    if verbose:
        st = controller.status()
        log.sim_footer(total_score, total_max,
                       st["inactive_count"], st["active_count"],
                       email_score=email_total_score,
                       email_max=email_total_max)
        log.type_breakdown(by_type, title="Score Breakdown by Type (per scenario)")
        if email_total_max:
            log.type_breakdown(by_type_email,
                               title="Score Breakdown by Type (per email)")

    if _model_harness is not None:
        _model_harness.shutdown()

    return {
        "total_score": total_score,
        "total_max": total_max,
        "email_score": email_total_score,
        "email_max": email_total_max,
        "daily_log": daily_log,
        "by_type": dict(by_type),
        "by_type_email": dict(by_type_email),
        "remaining_inactive": len(controller.inactive_pool),
        "remaining_active": len(controller.active_pool),
    }




# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else "Emails.xlsx"
    run_simulation(path, verbose=True)
