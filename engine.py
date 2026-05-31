from __future__ import annotations

import argparse
import calendar as _calendar
import os
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
from app.models.calendar import CalendarResponse


# Mirrors grader._NO_ACTION. Used to override the per-email grade when the AI
# took a destructive action (delete_*) that doesn't appear in the added/modified
# diff — the grader's len(items)==0 check would otherwise falsely pass.
_NO_ACTION_RE = re.compile(r"no\s*action", re.IGNORECASE)

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

# Weekday / ordinal / month lookups for the relative-date tokens the resolver
# understands. Names are matched after the token is lowercased + whitespace-
# stripped (see _resolve_one_token). Python weekday(): Monday=0 .. Sunday=6.
_WEEKDAYS = {"monday": 0, "tuesday": 1, "wednesday": 2, "thursday": 3,
             "friday": 4, "saturday": 5, "sunday": 6}
_WEEKDAY_ALT = "|".join(_WEEKDAYS)  # regex alternation: "monday|tuesday|..."
_ORDINALS = {"first": 1, "1st": 1, "second": 2, "2nd": 2, "third": 3, "3rd": 3,
             "fourth": 4, "4th": 4, "fifth": 5, "5th": 5, "last": -1}
_MONTHS = {name.lower(): idx for idx, name in enumerate(_calendar.month_name) if name}


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


def _next_weekday(sim_date: datetime, target_wd: int) -> datetime:
    """Next occurrence of weekday `target_wd` (Mon=0..Sun=6) on or after sim_date.

    "On or after" mirrors {date-Nth}: if sim_date already falls on target_wd the
    served day itself is returned (offset 0)."""
    return sim_date + timedelta(days=(target_wd - sim_date.weekday()) % 7)


def _next_week_weekday(sim_date: datetime, target_wd: int) -> datetime:
    """The `target_wd` weekday in the (Monday-based) week AFTER sim_date's week.

    "Next Wednesday" = the Wednesday of next week, not the next Wednesday a few
    days out. Anchoring on next week's Monday makes this consistent with the
    existing {nextweek-date} token (= sim_date + 7) for the same weekday."""
    this_monday = sim_date - timedelta(days=sim_date.weekday())
    return this_monday + timedelta(days=7 + target_wd)


def _nth_weekday_of_month(ref: datetime, year: int, month: int,
                          target_wd: int, n: int) -> Optional[datetime]:
    """Date of the n-th `target_wd` in (year, month); n == -1 means the last one.

    Returns a datetime at midnight (ref's tzinfo), or None when the month has no
    n-th occurrence (e.g. a 5th Friday in a 28-day month)."""
    first_wd, days_in_month = _calendar.monthrange(year, month)  # first_wd: Mon=0
    first_occ = 1 + (target_wd - first_wd) % 7  # day-of-month of the 1st target_wd
    if n == -1:
        dom = first_occ
        while dom + 7 <= days_in_month:
            dom += 7
    else:
        dom = first_occ + (n - 1) * 7
        if dom > days_in_month:
            return None
    return ref.replace(year=year, month=month, day=dom,
                       hour=0, minute=0, second=0, microsecond=0)


def _resolve_nth_weekday(sim_date: datetime, n: int, target_wd: int,
                         qualifier: Optional[str]) -> Optional[datetime]:
    """Resolve an Nth-weekday-of-month token to a concrete date.

    qualifier:
      None         -> next future occurrence (this month if its n-th weekday is
                      on/after sim_date, else roll forward), like {date-Nth}.
      "nextmonth"  -> the n-th weekday of the month after sim_date's month.
      "of<month>"  -> the n-th weekday of the named month's next future
                      occurrence (this year if still upcoming, else next year)."""
    base = sim_date.replace(hour=0, minute=0, second=0, microsecond=0)
    if qualifier == "nextmonth":
        nm = _add_months(base.replace(day=1), 1)
        return _nth_weekday_of_month(base, nm.year, nm.month, target_wd, n)
    if qualifier and qualifier.startswith("of"):
        month = _MONTHS.get(qualifier[2:])
        if not month:
            return None
        for year in (base.year, base.year + 1):
            cand = _nth_weekday_of_month(base, year, month, target_wd, n)
            if cand is not None and cand.date() >= sim_date.date():
                return cand
        return None
    # bare: roll forward month-by-month until an occurrence lands on/after today.
    year, month = base.year, base.month
    for _ in range(14):  # 14 > 12 so a missing 5th-weekday month can't trap us
        cand = _nth_weekday_of_month(base, year, month, target_wd, n)
        if cand is not None and cand.date() >= sim_date.date():
            return cand
        nm = _add_months(datetime(year, month, 1, tzinfo=base.tzinfo), 1)
        year, month = nm.year, nm.month
    return None


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

    # {date+N- TIME} -> N days out, at that time. E.g. {date+1- 3PM} ->
    # "...+1 day at 03:00 PM", {date+2- 11AM}. (Whitespace strip already
    # collapsed "date+1- 3PM" to "date+1-3pm".)
    m = re.match(r"^date\+(\d+)-(.+)$", normalized)
    if m:
        time_parsed = _parse_time(m.group(2))
        if time_parsed is not None:
            hour, minute = time_parsed
            return (sim_date + timedelta(days=int(m.group(1))))\
                .replace(hour=hour, minute=minute, second=0, microsecond=0)\
                .strftime(_DATETIME_FMT)

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

    # {nextweek-<weekday>} -> that weekday in next (Monday-based) week, with an
    # optional time and/or timezone label: {nextweek-wednesday}, {nextweek-friday},
    # {nextweek-Thursday, 3pm GMT}. A trailing "GMT"/"UTC" is dropped (the whole
    # simulation runs in UTC), so "3pm GMT" -> 15:00.
    m = re.match(r"^nextweek-?(" + _WEEKDAY_ALT + r")(?:,?(.+))?$", normalized)
    if m:
        base = _next_week_weekday(sim_date, _WEEKDAYS[m.group(1)])\
            .replace(hour=0, minute=0, second=0, microsecond=0)
        rest = m.group(2)
        if not rest:
            return base.strftime(_DATE_FMT)
        time_parsed = _parse_time(re.sub(r"(gmt|utc)$", "", rest))
        if time_parsed is None:
            return None  # weekday recognized but the trailing text isn't a time
        return base.replace(hour=time_parsed[0], minute=time_parsed[1])\
            .strftime(_DATETIME_FMT)

    # {date- this week TIME} -> today at that time. E.g. {date- this week 11pm}
    # normalizes to "date-thisweek11pm". A bare {date-thisweek} (no time) falls
    # through to the "today" rule just below.
    m = re.match(r"^date-?this-?week(.+)$", normalized)
    if m:
        time_parsed = _parse_time(re.sub(r"(gmt|utc)$", "", m.group(1)))
        if time_parsed is not None:
            hour, minute = time_parsed
            return sim_date.replace(hour=hour, minute=minute, second=0, microsecond=0)\
                .strftime(_DATETIME_FMT)

    # {date-thisweek} / {date-this-week} -> ambiguous, treat as today
    if re.match(r"^date-?this-?week$", normalized):
        return sim_date.strftime(_DATE_FMT)

    # {date-14th}, {date-25th} -> next future occurrence of that day-of-month
    m = re.match(r"^date-(\d{1,2})(?:st|nd|rd|th)$", normalized)
    if m:
        return _next_day_of_month(sim_date, int(m.group(1))).strftime(_DATE_FMT)

    # {date-<weekday>} -> next future occurrence of that weekday on/after today.
    # E.g. {date-Tuesday}. (Distinct from {nextweek-<weekday>}, which jumps a week.)
    m = re.match(r"^date-(" + _WEEKDAY_ALT + r")$", normalized)
    if m:
        return _next_weekday(sim_date, _WEEKDAYS[m.group(1)]).strftime(_DATE_FMT)

    # {date-10AM}, {date-1:14PM}, {date-12:30PM} -> today at that time
    m = re.match(r"^date-(.+)$", normalized)
    if m:
        time_parsed = _parse_time(m.group(1))
        if time_parsed is not None:
            hour, minute = time_parsed
            return sim_date.replace(hour=hour, minute=minute, second=0, microsecond=0)\
                .strftime(_DATETIME_FMT)

    # {third Friday}, {3rd Friday of November}, {third Friday nextmonth},
    # {third Friday -1 dinner} -> the n-th weekday of a month (default = next
    # future occurrence; "nextmonth" or "of <month>" override the month), plus an
    # optional +/-N day offset. Any trailing label ("dinner") is ignored.
    m = re.match(
        r"^(first|second|third|fourth|fifth|last|[1-5](?:st|nd|rd|th))"
        r"(" + _WEEKDAY_ALT + r")"
        r"(nextmonth|of[a-z]+)?"
        r"([+-]\d+)?.*$",
        normalized,
    )
    if m:
        result = _resolve_nth_weekday(sim_date, _ORDINALS[m.group(1)],
                                      _WEEKDAYS[m.group(2)], m.group(3))
        if result is None:
            return None
        if m.group(4):
            result += timedelta(days=int(m.group(4)))
        return result.strftime(_DATE_FMT)

    return None


def resolve_tokens(text: str, sim_date: datetime, keep_braces: bool = False,
                   skip_bare_date: bool = False) -> str:
    """Replace every {...} date/link token in `text` with a resolved value.

    Tokens that don't match a known pattern (e.g. `{Annual Report 1}`,
    `{contract 2}`, range tokens like `{date-12:30-2:00PM}`) are left as-is
    so they remain visible for downstream review.

    keep_braces controls whether the braces survive resolution:
    - False (default, for subject/body): `{date}` -> `March 15, 2000`.
    - True (for success_criteria): `{date}` -> `{March 15, 2000}`. The grader
      reads criteria dates from *inside* braces (grader._extract_date_token),
      and keeping the braces also stops `grader._parse_sub_criteria` from
      splitting on the comma inside a resolved date like "March 15, 2000".
      Unresolved tokens keep their original braces either way, so the grader
      still skips them (FIX-2).

    skip_bare_date (criteria only): leave a *bare* `{date}` token UNRESOLVED so
    the grader treats it as "no specific date" and checks existence only, while
    specific tokens (`{date+2}`, `{date-14th}`, `{date-3PM}`, `{date-nextweek}`)
    still resolve and get strict date-matching. This is deliberate: the dataset
    (Emails.xlsx) uses bare `CC-{date}`/`TC-{date}` when the email's real target
    date can't be expressed as a token. Resolving bare `{date}` to the arbitrary
    served day and strict-matching it would FALSE-NEGATIVE a correct model.

    Note: the live path (apply_date_substitutions / resolve_scenario_criteria)
    does NOT pass this flag — bare `{date}` resolves to the served day and is
    strict-matched, per the decision in HANDOFF.md §4 (the answer key is fixed by
    hand where the served day is wrong). See docs/TOKEN_REFERENCE.md for the
    token forms and docs/REMAINING_WORK.md (G1/G5) for the rationale.
    """
    if not text:
        return text

    def _sub(match: re.Match) -> str:
        inner = match.group(1)
        if skip_bare_date and re.sub(r"\s+", "", inner.lower()) == "date":
            return match.group(0)  # leave bare {date} unresolved -> existence-only
        resolved = _resolve_one_token(inner, sim_date)
        if resolved is None:
            return match.group(0)
        return "{" + resolved + "}" if keep_braces else resolved

    return _TOKEN_RE.sub(_sub, text)


def apply_date_substitutions(email: Email, sim_date: datetime) -> Email:
    """Return a copy of `email` with subject, body, AND success_criteria tokens
    resolved at `sim_date`.

    Resolving success_criteria here is what makes the grader's date matching
    actually fire (FIX-2): criteria like `CC-{date}` reach the grader as a
    concrete date instead of a raw token, which `grader._extract_date_token`
    skips. This fixes the per-email grading path; the per-scenario path uses
    `resolve_scenario_criteria` below (each criterion at its own email's date).
    """
    e = deepcopy(email)
    e.subject = resolve_tokens(e.subject, sim_date)
    e.body = resolve_tokens(e.body, sim_date)
    if e.success_criteria:
        e.success_criteria = resolve_tokens(e.success_criteria, sim_date, keep_braces=True)
    return e


def resolve_scenario_criteria(
    scenario: Scenario, controller: FlowController, fallback_date: datetime
) -> Scenario:
    """Return a copy of `scenario` whose success_criteria are resolved, each at
    the sim_date the email that defined it was actually served (FIX-2).

    Scenario-level criteria are collected from individual emails (loader order =
    chain order of emails that carry criteria), and a chain's emails can be
    served on different days. Resolving every criterion at completion day would
    be wrong, so we look up each email's served date in the controller's
    delivery_log and resolve against that. Emails not found in the log (should
    not happen for a completed scenario) fall back to `fallback_date`.
    """
    served: dict[int, datetime] = {}
    for entry in controller.delivery_log:
        if entry.get("scenario_id") != scenario.scenario_id:
            continue
        raw = entry.get("sim_date")
        if not raw:
            continue
        try:
            served[entry["email_index"]] = datetime.strptime(raw, "%Y-%m-%d").replace(tzinfo=timezone.utc)
        except (ValueError, TypeError):
            pass

    resolved_criteria = [
        resolve_tokens(email.success_criteria, served.get(i, fallback_date), keep_braces=True)
        for i, email in enumerate(scenario.emails)
        if email.success_criteria
    ]

    s = deepcopy(scenario)
    s.success_criteria = resolved_criteria
    return s


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
    scenarios=None,
    grade_by_scenario: bool = True,
    harness: str = "claude-code",
    model: str = "claude-haiku-4-5",
    reasoning: Optional[str] = None,
    api_base: Optional[str] = None,
    openrouter_key: Optional[str] = None,
    conversation_continuity: Optional[bool] = None,
) -> dict:
    """Run the full N-day benchmark simulation and return aggregated results.

    conversation_continuity: None defers to the CONVERSATION_CONTINUITY env var
    (read inside the adapter). Pass True/False to force it on the live path —
    this is what makes `CONVERSATION_CONTINUITY=0 python engine.py` and the A/B
    knob actually reach the harness (FIX-8).
    """
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

    # FIX-9: compaction / token dimension. Lets the score sheet tell
    # "succeeded within the context window" from "succeeded through compaction."
    total_input_tokens = 0
    total_output_tokens = 0
    total_compactions = 0
    total_ungraded = 0
    compaction_split: dict[str, dict[str, int]] = {
        "within_window": {"count": 0, "score": 0, "max_score": 0},
        "through_compaction": {"count": 0, "score": 0, "max_score": 0},
    }

    # FIX-6: optional one-shot retry for transient harness failures. Off by
    # default — a benchmark run should record failures, not paper over them.
    retry_on_failure = os.environ.get("HARNESS_RETRY_ON_FAILURE", "0") == "1"

    adapter = None
    if model_fn is None:
        try:
            from harness import get_adapter
            adapter = get_adapter(
                harness,
                model=model,
                reasoning=reasoning,
                api_base=api_base,
                openrouter_key=openrouter_key,
                conversation_continuity=conversation_continuity,
            )
        except ImportError:
            pass

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
            if adapter is not None:
                adapter.start_session(scenario_str_to_int(activated.scenario_id))

        if ready and verbose:
            log.day_header(day + 1, sim_date.strftime("%Y-%m-%d"), len(ready))

        for email, scenario, idx in ready:
            resolved = apply_date_substitutions(email, sim_date)

            # Snapshot state BEFORE the model acts, so we can diff and attribute
            # what changed to this email specifically (not to earlier emails
            # in the chain).
            state_before = fetch_scenario_results(scenario)

            # FIX-6: a crashed/timed-out turn must NOT silently vanish. We
            # capture the error, still grade the email against the diff (an
            # attempted-but-failed turn scores whatever it managed — usually 0),
            # record the error in the delivery log, and let the chain continue.
            # The adapter resets its session on failure so the next email starts
            # fresh instead of resuming a dead one.
            turn_error: Optional[str] = None
            sid_int = scenario_str_to_int(scenario.scenario_id)
            if model_fn is not None:
                model_fn(resolved, sim_date)
            elif adapter is not None:
                attempts = 2 if retry_on_failure else 1
                for attempt in range(attempts):
                    try:
                        adapter.run_turn(resolved, sim_date, scenario_id=sid_int)
                        turn_error = None
                        break
                    except Exception as exc:
                        turn_error = str(exc)
                        import bench_logger as log
                        log.error("engine",
                                  f"adapter.run_turn failed on scenario {scenario.scenario_id} "
                                  f"(attempt {attempt + 1}/{attempts}): {exc}")
            elif _HAS_MODEL_RUNNER:
                try:
                    run_model_turn(resolved, sim_date, scenario_id=sid_int)
                except Exception as exc:
                    turn_error = str(exc)
                    import bench_logger as log
                    log.error("engine",
                              f"run_model_turn failed on scenario {scenario.scenario_id}: {exc}")
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
                error=turn_error,
            )

        # Grade scenarios that completed today (per-scenario, matches grader design)
        day_score = 0
        day_max = 0
        for scenario in controller.completed_scenarios_today():
            state = fetch_scenario_results(scenario)
            # Resolve each criterion at the sim_date its email was served (FIX-2)
            # before grading, so date checks like CC-{date} actually match.
            graded_scenario = resolve_scenario_criteria(scenario, controller, sim_date)
            result = define_grading_system(graded_scenario, state["calendar"], state["todos"],
                                          grade_by_scenario=grade_by_scenario)

            # FIX-9: read per-scenario telemetry BEFORE end_session releases it,
            # and attach the compaction/token dimension to the result.
            sid_int = scenario_str_to_int(scenario.scenario_id)
            tel = adapter.scenario_telemetry(sid_int) if adapter is not None else {}
            compactions = tel.get("compactions", 0)
            compacted = compactions > 0
            result["compaction_triggered"] = compacted
            # Reserved: claude compacts before overflowing, so an honest "exceeded"
            # signal only exists if a turn errored on context — left False here.
            result["context_window_exceeded"] = False
            result["input_tokens"] = tel.get("input_tokens", 0)
            result["output_tokens"] = tel.get("output_tokens", 0)
            total_input_tokens += result["input_tokens"]
            total_output_tokens += result["output_tokens"]
            total_compactions += compactions
            total_ungraded += len(result.get("ungraded", []))
            grp = compaction_split["through_compaction" if compacted else "within_window"]
            grp["count"] += 1
            grp["score"] += result["score"]
            grp["max_score"] += result["max_score"]

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
            if adapter is not None:
                adapter.end_session(scenario_str_to_int(scenario.scenario_id))

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
        log.compaction_breakdown(compaction_split, total_compactions,
                                 total_input_tokens, total_output_tokens,
                                 ungraded=total_ungraded)

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
        # FIX-9: compaction / token dimension.
        "compaction_scenarios": compaction_split["through_compaction"]["count"],
        "within_window_scenarios": compaction_split["within_window"]["count"],
        "total_compactions": total_compactions,
        "total_input_tokens": total_input_tokens,
        "total_output_tokens": total_output_tokens,
        "compaction_split": compaction_split,
        "ungraded_criteria": total_ungraded,
    }




# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Run SecretaryBench simulation.")
    parser.add_argument("path", nargs="?", default="Emails.xlsx")
    parser.add_argument("--harness", default="claude-code")
    # --model / --reasoning default to the CLAUDE_MODEL / CLAUDE_REASONING env
    # vars so both the flag and the env knob reach the live path (FIX-8/FIX-11).
    parser.add_argument("--model", default=os.environ.get("CLAUDE_MODEL", "claude-haiku-4-5"))
    parser.add_argument("--reasoning", choices=["low", "medium", "high"],
                        default=os.environ.get("CLAUDE_REASONING") or None)
    parser.add_argument("--openrouter", action="store_true")
    parser.add_argument("--api-base", default=None, dest="api_base")
    parser.add_argument(
        "--continuity", choices=["0", "1"], default=None,
        help="Override CONVERSATION_CONTINUITY: 1 resumes per-scenario sessions, "
             "0 runs a fresh session per email. Defaults to the env var.",
    )
    parser.add_argument("--days", type=int, default=SIM_DAYS,
                        help=f"Number of simulated days to run (default {SIM_DAYS}).")
    parser.add_argument("--limit", type=int, default=None,
                        help="Run only the first N scenarios (bounded smoke test). "
                             "Note: the scheduler spreads scenarios across --days, so "
                             "--days alone does not reduce total turns; --limit does.")
    args = parser.parse_args()

    subset = None
    if args.limit is not None:
        subset = load_scenarios(args.path)[: args.limit]

    run_simulation(
        args.path,
        sim_days=args.days,
        scenarios=subset,
        verbose=True,
        harness=args.harness,
        model=args.model,
        reasoning=args.reasoning,
        api_base=args.api_base,
        openrouter_key=os.environ.get("OPENROUTER_API_KEY") if args.openrouter else None,
        conversation_continuity=(None if args.continuity is None else args.continuity == "1"),
    )
