from __future__ import annotations

import calendar as _calendar
import re
import sys
import time
from collections import defaultdict
from copy import deepcopy
from datetime import datetime, timedelta, timezone
from typing import Optional

from loader import load_scenarios, Email, Scenario
from flow_controller import FlowController
from grader import define_grading_system
from pipeline import register_scenario, fetch_scenario_results, scenario_str_to_int

# Bridge to Person 3's model runner. The agreed-on signature is
#   run_model_turn(email: Email, sim_date: datetime) -> None
# (one LLM turn per email, with the MCP tools attached on the runner side).
# Falls back to the in-file mock if model_runner.py hasn't been added yet so
# the simulation still runs end-to-end before Miguel's piece lands.
try:
    from model_runner import run_model_turn, scenario_completed  # type: ignore
    _HAS_MODEL_RUNNER = True
except ImportError:
    run_model_turn = None  # type: ignore
    scenario_completed = None  # type: ignore
    _HAS_MODEL_RUNNER = False


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


def run_simulation(
    path: str = "Emails.xlsx",
    sim_days: int = SIM_DAYS,
    sim_start: datetime = SIM_START,
    seed: Optional[int] = 42,
    verbose: bool = True,
    model_fn=None,
    scenarios=None,
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
    daily_log: list[dict] = []
    # Per-scenario-type breakdown so we can see which categories the model
    # handles well vs poorly. Same shape across model swaps, so a Haiku vs
    # Sonnet comparison is just diffing two of these.
    by_type: dict[str, dict[str, int]] = defaultdict(
        lambda: {"count": 0, "score": 0, "max_score": 0}
    )


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
            if model_fn is not None:
                model_fn(resolved, sim_date)
            elif _HAS_MODEL_RUNNER:
                run_model_turn(resolved, sim_date,
                               scenario_id=scenario_str_to_int(scenario.scenario_id))
            else:
                model_interaction_mock([resolved], sim_date, verbose=verbose)
            controller.mark_served(scenario.scenario_id, idx)

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
            # Free the persistent model conversation for this scenario — keeps
            # the runner's per-scenario dict from growing unbounded over a run.
            if scenario_completed is not None:
                scenario_completed(scenario_str_to_int(scenario.scenario_id))

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
                       st["inactive_count"], st["active_count"])
        log.type_breakdown(by_type)

    return {
        "total_score": total_score,
        "total_max": total_max,
        "daily_log": daily_log,
        "by_type": dict(by_type),
        "remaining_inactive": len(controller.inactive_pool),
        "remaining_active": len(controller.active_pool),
    }




# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else "Emails.xlsx"
    run_simulation(path, verbose=True)
