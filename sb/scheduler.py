"""
sb.scheduler — turns the email DAG into a serve plan (SERVING_AND_SCHEMA.md §4).

Key property: the serve plan does NOT depend on the model. Readiness depends only
on which emails are served and on anchor values, and anchor values come from
emission expressions evaluated at the emitter's serve day — all deterministic. So
we compute the entire plan (and check feasibility) BEFORE any model runs.

Policy: seeded EDF + weighted random filler.
  - EDF: any ready email whose serve-by deadline is within URGENCY_HORIZON is
    forced first, earliest deadline first — so a `date` email is never stranded.
  - Filler: sample the rest to hit a per-day target (1..DAILY_MAX). This is the
    difficulty/ambiguity knob.

Same (corpus, start_date, seed, levers) => identical plan, 100% reproducible.
"""
from __future__ import annotations

import random
from dataclasses import dataclass, field
from datetime import date, timedelta

from sb import resolver
from sb.resolver import Context, Value
from sb.schema import Corpus


class InfeasibleSchedule(RuntimeError):
    """Raised when the corpus cannot be served within the constraints."""


@dataclass
class Levers:
    daily_min: int = 1
    daily_max: int = 5
    urgency_horizon: int = 7          # days before a deadline we start forcing the email


@dataclass
class Plan:
    start_date: date
    per_day: list[list[str]]                       # day index -> email ids served
    serve_date: dict[str, date]                    # email id -> serve date
    anchors: dict[str, Value]                      # final anchor table (stable, superset-safe)
    deadlines: dict[str, date] = field(default_factory=dict)

    @property
    def n_days(self) -> int:
        return len(self.per_day)


def _anchor_exprs(corpus: Corpus, email_id: str) -> list[str]:
    """Answer predicate expressions of an email that reference an anchor."""
    from sb.schema import _email_predicates, _refs_in
    return [e for e in _email_predicates(corpus.emails[email_id].answer) if _refs_in(e)]


def build_plan(
    corpus: Corpus,
    *,
    start_date: date,
    seed: int,
    n_days: int = 100,
    levers: Levers | None = None,
) -> Plan:
    levers = levers or Levers()
    rng = random.Random(seed)
    emails = corpus.emails

    served: dict[str, int] = {}
    serve_date: dict[str, date] = {}
    anchors: dict[str, Value] = {}
    deadlines: dict[str, date] = {}
    per_day: list[list[str]] = []

    def has_date_edge(eid: str) -> bool:
        return any(d.type == "date" for d in emails[eid].depends_on)

    def update_deadlines(today: date) -> None:
        for eid, email in emails.items():
            if eid in served or eid in deadlines or not has_date_edge(eid):
                continue
            if not email.anchor_refs.issubset(anchors.keys()):
                continue
            ctx = Context(today, anchors)
            dates = [resolver._as_date(resolver.resolve(expr, ctx)) for expr in _anchor_exprs(corpus, eid)]
            if dates:
                deadlines[eid] = min(dates)

    def is_ready(eid: str, day: int) -> bool:
        email = emails[eid]
        for dep in email.depends_on:
            if served.get(dep.email) is None or served[dep.email] >= day:
                return False
        if has_date_edge(eid):
            dl = deadlines.get(eid)
            if dl is None or (start_date + timedelta(days=day)) > dl:
                return False
        return True

    for day in range(n_days):
        today = start_date + timedelta(days=day)
        update_deadlines(today)

        if len(served) == len(emails):
            break

        ready = [eid for eid in corpus.topo_order() if eid not in served and is_ready(eid, day)]
        if not ready:
            per_day.append([])
            continue

        horizon_cut = today + timedelta(days=levers.urgency_horizon)
        urgent = sorted([e for e in ready if e in deadlines and deadlines[e] <= horizon_cut],
                        key=lambda e: deadlines[e])
        rest = [e for e in ready if e not in urgent]
        rng.shuffle(rest)

        target = rng.randint(levers.daily_min, levers.daily_max)
        fill = rest[: max(0, target - len(urgent))]
        batch = (urgent + fill)[: levers.daily_max]

        for eid in batch:
            served[eid] = day
            serve_date[eid] = today
            ctx = Context(today, anchors)
            for name, expr in emails[eid].emits.items():
                anchors[name] = resolver.resolve(expr, ctx)
        per_day.append(batch)

    # --- feasibility ---
    missing = set(emails) - set(served)
    if missing:
        raise InfeasibleSchedule(
            f"{len(missing)} email(s) never served within {n_days} days: {sorted(missing)} "
            f"(over-constrained windows or n_days too small)"
        )
    for eid, dl in deadlines.items():
        if serve_date[eid] > dl:
            raise InfeasibleSchedule(
                f"email {eid!r} served {serve_date[eid]} after its serve-by deadline {dl}"
            )

    return Plan(start_date=start_date, per_day=per_day, serve_date=serve_date,
                anchors=anchors, deadlines=deadlines)
