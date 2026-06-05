"""
sb.probe — synthesize a multi-storyline corpus to PROBE cross-storyline calendar time
conflicts. A measurement fixture (sibling of sb.scale), not part of the graded corpus.

The real ~20-author corpus lives in the prod store; until it's pulled (python -m sb.sync),
this stands in for it with the shape that drives the conflicts question: many independent
storylines, each booking a few timed events at the clock times authors actually favor
(9-10, 10-11, 11-12, 1-2, 2-3, 3-4). Everything is serve-relative (next:WD @HH:MM-HH:MM)
so there are no anchors to collide and each storyline oracle-solves on its own. The
scheduler interleaves them on one calendar; sb.conflicts then counts the overlaps.

    python -m sb.probe --authors 20 --events 3 --seed 42 --dst build/probe
    python -m sb.conflicts --corpus build/probe --seed 42 --days 200

Once the real corpus is pulled, run sb.conflicts straight on it — that's the number the
resample-vs-serving-constraint decision should use; this probe is only the offline stand-in.
"""
from __future__ import annotations

import argparse
import json
import random
import shutil
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]

# Clock slots a human actually picks for a meeting (each one hour long). Weighted toward
# mid-morning and early afternoon, the way real calendars cluster.
FAVES = ["09:00-10:00", "10:00-11:00", "11:00-12:00", "13:00-14:00", "14:00-15:00", "15:00-16:00"]
FAVE_WEIGHTS = [5, 4, 2, 2, 5, 3]
WEEKDAYS = ["MON", "TUE", "WED", "THU", "FRI"]
TOPICS = ["sync", "review", "standup", "1on1", "planning", "demo", "retro", "interview", "checkin", "briefing"]


def _author_node(idx: int, n_events: int, rng: random.Random, spread_weeks: int, slot_pool: list[str]) -> dict:
    nid = f"author{idx:02d}"
    emails = []
    weights = FAVE_WEIGHTS if slot_pool is FAVES else None
    for j in range(n_events):
        wd = rng.choice(WEEKDAYS)
        slot = rng.choices(slot_pool, weights=weights, k=1)[0]
        # Spread events into the future like real needles do (+0..spread_weeks weeks), so they don't
        # all bunch into the week right after the email serves. Offset rides directly on the base.
        wk = rng.randint(0, spread_weeks)
        expr = f"next:{wd}{f'+{wk}w' if wk else ''} @{slot}"
        topic = rng.choice(TOPICS)
        name = f"{topic}{j+1}"
        emails.append({
            "id": f"{nid}.e{j+1}",
            "from": "COLLEAGUE",
            "to": "CEO",
            "subject": f"{topic.title()} request #{j+1}",
            "body": f"Can we put the {topic} on the calendar for next:{wd} at {slot.replace('-', ' to ')}?",
            "depends_on": [],
            "answer": {"ops": [{"create": name, "kind": "event", "on": {"eq": expr}}]},
        })
    return {"id": nid, "cast": {"CEO": "you", "COLLEAGUE": "A Colleague"}, "emails": emails}


def build(dst: Path, authors: int, events: int, seed: int, spread_weeks: int = 0, slots: int = len(FAVES)) -> None:
    nodes = dst / "nodes"
    if dst.exists():
        shutil.rmtree(dst)
    nodes.mkdir(parents=True)
    rng = random.Random(seed)
    # <=6 -> the favorites (weighted); >6 -> a wider, evenly-picked pool (authors with more varied times).
    pool = FAVES if slots <= len(FAVES) else _wider_pool(slots)
    for i in range(1, authors + 1):
        node = _author_node(i, events, rng, spread_weeks, pool)
        (nodes / f"{node['id']}.json").write_text(json.dumps(node, indent=2))


def _wider_pool(n: int) -> list[str]:
    """A wider clock pool of one-hour slots starting every half hour from 08:00 — for modeling
    authors who DON'T all pick the same handful of times (the 'more varied times' case)."""
    starts = [(h, m) for h in range(8, 17) for m in (0, 30)]            # 08:00 .. 16:30
    slots = [f"{h:02d}:{m:02d}-{h+1:02d}:{m:02d}" for h, m in starts]   # each one hour long
    return slots[:n]


def main() -> None:
    ap = argparse.ArgumentParser(description="synthesize a probe corpus for sb.conflicts")
    ap.add_argument("--authors", type=int, default=20)
    ap.add_argument("--events", type=int, default=3)
    ap.add_argument("--seed", type=int, default=42)
    ap.add_argument("--spread-weeks", type=int, default=0, help="scatter events 0..N weeks out (like real needles)")
    ap.add_argument("--slots", type=int, default=len(FAVES), help="size of the clock-time pool (bigger = more varied times)")
    ap.add_argument("--dst", default="build/probe")
    a = ap.parse_args()
    dst = REPO / a.dst
    build(dst, a.authors, a.events, a.seed, a.spread_weeks, a.slots)
    print(f"wrote {a.authors} storylines x {a.events} timed events "
          f"(spread {a.spread_weeks}w, {a.slots} slots) -> {dst}")


if __name__ == "__main__":
    main()
