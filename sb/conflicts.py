"""
sb.conflicts — monitor cross-storyline TIME conflicts on the shared calendar.

Grading is node-scoped, so two storylines whose events land on the same clock slot do NOT
affect each other's score — a clash here is cosmetic. This is the **monitor-only** tracking
that BACKLOG.md §2 ("Global calendar + global grading") refers to: handling conflicts is
deferred to that global-grading epic (see §2a), since under scoped grading there is nothing
to fix. This tool just produces the number; it changes nothing about grading or serving.

Run it on the real corpus when picking that epic up:  python -m sb.conflicts --corpus corpus

It resolves every served event to its FINAL booked slot, replaying create/move/cancel
per obligation exactly as the grader reconciles node state (a move supersedes the
create's slot; a cancel vacates it), then counts overlapping timed slots that belong to
DIFFERENT storylines. Only timed events (an @HH:MM-HH:MM span, or an @HH:MM point) hold a
clock slot; bare-day events and ambiguous predicates (any_of / by) have no single slot and
are reported separately, not as conflicts.

    python -m sb.conflicts --corpus build/scaled --seed 42 --days 200
"""
from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import date, datetime

from sb import resolver
from sb.resolver import Context, TimeInterval, Value
from sb.schema import Corpus, load_corpus
from sb.scheduler import build_plan


@dataclass(frozen=True)
class Slot:
    node: str
    email_id: str
    name: str            # obligation name (its title-ish label)
    span: TimeInterval   # the booked [start, end) on the shared calendar


def _slot_span(val: Value) -> TimeInterval | None:
    """The clock slot an event occupies, or None for a day-level value (no fixed time)."""
    if isinstance(val, TimeInterval):
        return val
    if isinstance(val, datetime):
        return TimeInterval(val, val)         # a point time @HH:MM (zero-length)
    return None                                # bare date / week / month -> no clock slot


def _overlap(a: TimeInterval, b: TimeInterval) -> bool:
    # Spans use the half-open rule (touching edges don't conflict). A point (zero-length)
    # conflicts if it falls inside the other span or coincides with another point.
    if a.start == a.end or b.start == b.end:
        return max(a.start, b.start) <= min(a.end, b.end)
    return a.overlaps(b)


@dataclass
class ConflictReport:
    served_emails: int
    event_ops: int          # create/move event ops seen (any predicate)
    timed_slots: int        # final surviving events that hold a clock slot
    day_level: int          # final surviving events with no clock slot (bare day)
    ambiguous: int          # event ops whose predicate is any_of / by (no single slot)
    conflicts: list[tuple[Slot, Slot]]
    peak_concurrency: int   # most events booked over a single instant, anywhere

    @property
    def conflicting_slots(self) -> int:
        s: set[tuple[str, str]] = set()
        for a, b in self.conflicts:
            s.add((a.node, a.name)); s.add((b.node, b.name))
        return len(s)

    @property
    def conflict_rate(self) -> float:
        return self.conflicting_slots / self.timed_slots if self.timed_slots else 0.0


def _serve_order(plan) -> list[str]:
    return [eid for batch in plan.per_day for eid in batch]


def _final_slots(corpus: Corpus, plan) -> tuple[list[Slot], dict]:
    """Replay served ops in serve order, keeping the surviving timed event slot per
    (node, obligation). Returns the live slots plus the running counts."""
    live: dict[tuple[str, str], Slot] = {}
    day_level: set[tuple[str, str]] = set()
    counts = {"event_ops": 0, "ambiguous": 0}
    for eid in _serve_order(plan):
        email = corpus.emails[eid]
        ctx = Context(plan.serve_date[eid], plan.anchors)
        for op in email.answer.ops:
            if op.kind != "event":
                continue                                   # only events occupy the calendar
            key = (email.node, op.name)
            if op.verb == "cancel":
                live.pop(key, None); day_level.discard(key)
                continue
            counts["event_ops"] += 1
            pred = op.on or {}
            if "eq" not in pred:                            # any_of / by: no single slot
                counts["ambiguous"] += 1
                continue
            span = _slot_span(resolver.resolve(pred["eq"], ctx))
            if span is None:
                live.pop(key, None); day_level.add(key)     # day-level event (no clock)
            else:
                day_level.discard(key)
                live[key] = Slot(email.node, eid, op.name, span)
    counts["day_level"] = len(day_level)
    return list(live.values()), counts


def find_conflicts(slots: list[Slot]) -> list[tuple[Slot, Slot]]:
    """Cross-storyline overlapping timed slots. O(n log n): sort by start, sweep."""
    out: list[tuple[Slot, Slot]] = []
    ordered = sorted(slots, key=lambda s: s.span.start)
    active: list[Slot] = []
    for s in ordered:
        active = [a for a in active if a.span.end > s.span.start or a.span.end == a.span.start]
        for a in active:
            if a.node != s.node and _overlap(a.span, s.span):
                out.append((a, s))
        active.append(s)
    return out


def _peak_concurrency(slots: list[Slot]) -> int:
    events: list[tuple[datetime, int]] = []
    for s in slots:
        events.append((s.span.start, 1))
        events.append((s.span.end if s.span.end > s.span.start else s.span.start, -1))
    # +1 before -1 at the same instant would over-count a touching edge; sort -1 first.
    events.sort(key=lambda e: (e[0], e[1]))
    cur = peak = 0
    for _, d in events:
        cur += d
        peak = max(peak, cur)
    return peak


def report(corpus: Corpus, plan) -> ConflictReport:
    slots, counts = _final_slots(corpus, plan)
    conflicts = find_conflicts(slots)
    return ConflictReport(
        served_emails=len(_serve_order(plan)),
        event_ops=counts["event_ops"],
        timed_slots=len(slots),
        day_level=counts["day_level"],
        ambiguous=counts["ambiguous"],
        conflicts=conflicts,
        peak_concurrency=_peak_concurrency(slots),
    )


def _fmt(s: Slot) -> str:
    a, b = s.span.start, s.span.end
    when = a.strftime("%a %b %d %H:%M")
    tail = f"-{b.strftime('%H:%M')}" if b > a else ""
    return f"{s.node}/{s.name} @ {when}{tail}"


def print_report(r: ConflictReport, show: int = 12) -> None:
    print(f"\nserved emails:        {r.served_emails}")
    print(f"event ops (create+move): {r.event_ops}")
    print(f"  final timed slots:  {r.timed_slots}  (these hold a clock slot)")
    print(f"  day-level events:   {r.day_level}  (bare day, no clock — never a time conflict)")
    print(f"  ambiguous (any_of/by): {r.ambiguous}  (no single slot — excluded)")
    print(f"\npeak concurrency:     {r.peak_concurrency} events booked over one instant")
    print(f"cross-storyline conflicts: {len(r.conflicts)} pair(s)")
    print(f"  slots in a conflict: {r.conflicting_slots} / {r.timed_slots}"
          f"  = {r.conflict_rate:.1%} of timed events")
    if r.conflicts:
        print("\n  examples:")
        for a, b in r.conflicts[:show]:
            print(f"    {_fmt(a)}\n      vs {_fmt(b)}")
        if len(r.conflicts) > show:
            print(f"    … and {len(r.conflicts) - show} more")
    print()


def main() -> None:
    ap = argparse.ArgumentParser(description="measure cross-storyline calendar time conflicts")
    ap.add_argument("--corpus", default="build/scaled")
    ap.add_argument("--seed", type=int, default=42)
    ap.add_argument("--days", type=int, default=200)
    ap.add_argument("--start", default="2026-06-01")
    a = ap.parse_args()
    corpus = load_corpus(a.corpus)
    plan = build_plan(corpus, start_date=date.fromisoformat(a.start), seed=a.seed, n_days=a.days)
    print_report(report(corpus, plan))


if __name__ == "__main__":
    main()
