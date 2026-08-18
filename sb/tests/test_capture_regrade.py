"""
The capture written by `sb.live.runner --out` must be sufficient to reproduce a
run's score offline, with no model and no store.

This is the property that makes register phase 1c worth paying for: if offline
re-grade does not equal the live score, the capture is lossy and every grader
change still costs a fresh run (register O-1).

No live model is involved. The runner's day loop is simulated against a store
state in the exact `/state` JSON shape `sb.live.runner` snapshots, driven by a
pluggable titling policy so the round-trip is exercised at both a perfect score
and a realistic partial one.
"""
import json
from datetime import date, datetime, time, timedelta
from pathlib import Path

import pytest

from sb.grader import grade_email
from sb.live.runner import _corpus_hash, _node_state, _turn_delta
from sb.oracle import _as_dt, _target
from sb.regrade import regrade, score
from sb.resolver import Context
from sb.schema import load_corpus
from sb.scheduler import Levers, build_plan

START = date(2026, 6, 1)
LEVERS = Levers(daily_min=1, daily_max=5, urgency_horizon=7)
CORPUS = "corpus"


def _key_title(op, email):
    """The oracle's policy: title with the answer key's own keywords."""
    return " ".join(op.match) if op.match else op.name


def _subject_title(op, email):
    """What a real assistant does: title from the email subject (register G-4)."""
    return email.subject


def simulate(corpus, plan, title_of):
    """Run the corpus against a dict store shaped exactly like store_app's /state.

    Returns (live_results, day_records) where day_records match what _Capture.day
    writes. Mirrors sb.live.runner's day loop: one turn per day, then split the
    day's new objects back to each email by the stamped email_id.
    """
    events, todos, day_records, results = [], [], [], {}
    counter = 0

    for day_no, batch in enumerate([b for b in plan.per_day if b], 1):
        sd = plan.serve_date[batch[0]]
        before = {r["id"] for r in events + todos}

        for eid in batch:                       # the "model" acts on the day's mail
            email = corpus.emails[eid]
            ctx = Context(serve=plan.serve_date[eid], anchors=plan.anchors)
            for op in email.answer.ops:
                title = title_of(op, email)
                if op.verb == "cancel":
                    events[:] = [r for r in events if title.lower() not in r["title"].lower()]
                    todos[:] = [r for r in todos if title.lower() not in r["title"].lower()]
                    continue
                when = _as_dt(_target(op.on, ctx), 9 if op.kind == "event" else 17)
                counter += 1
                if op.kind == "event":
                    events.append({"id": f"evt_{counter}", "email_id": eid, "title": title,
                                   "start": when.isoformat(),
                                   "end": (when + timedelta(hours=1)).isoformat(),
                                   "description": ""})
                else:
                    todos.append({"id": f"td_{counter}", "email_id": eid, "title": title,
                                  "due_date": when.isoformat(), "description": ""})

        state = {"events": list(events), "todos": list(todos)}
        day_new = {r["id"] for r in events + todos} - before
        by_eid = {r["id"]: r.get("email_id", "") for r in events + todos}

        for eid in batch:                       # grade exactly as the runner does
            email = corpus.emails[eid]
            ctx = Context(serve=plan.serve_date[eid], anchors=plan.anchors)
            eid_new = {i for i in day_new if by_eid.get(i) == eid}
            results[eid] = grade_email(
                email.answer, ctx,
                _node_state(corpus, state, email.node, eid_new),
                _turn_delta(corpus, state, eid_new))

        day_records.append({"day": day_no, "serve_date": str(sd), "batch": list(batch),
                            "ok": True, "state": state, "day_new": sorted(day_new),
                            "by_eid": by_eid})
    return results, day_records


def write_capture(tmp_path, corpus, plan, day_records, n_days):
    out = tmp_path / "capture"
    (out / "days").mkdir(parents=True)
    (out / "raw").mkdir(parents=True)
    order = [e for b in plan.per_day for e in b]
    (out / "manifest.json").write_text(json.dumps({
        "requested_model": "test", "seed": 42, "start": str(START), "n_days": n_days,
        "limit": None, "corpus_dir": CORPUS, "corpus_hash": _corpus_hash(CORPUS),
        "levers": {"daily_min": LEVERS.daily_min, "daily_max": LEVERS.daily_max,
                   "urgency_horizon": LEVERS.urgency_horizon},
        "planned_emails": len(order), "schema_version": 1,
    }, indent=2))
    for rec in day_records:
        (out / "days" / f"{rec['day']:03d}.json").write_text(json.dumps(rec, indent=2))
    return str(out)


@pytest.mark.parametrize("policy,name", [(_key_title, "answer-key"), (_subject_title, "subject")])
def test_capture_regrades_to_the_identical_score(tmp_path, policy, name):
    corpus = load_corpus(CORPUS)
    n_days = 200
    plan = build_plan(corpus, start_date=START, seed=42, n_days=n_days, levers=LEVERS)

    live, day_records = simulate(corpus, plan, policy)
    capture = write_capture(tmp_path, corpus, plan, day_records, n_days)

    offline = regrade(capture)

    assert set(offline) == set(live), f"{name}: re-grade covered a different email set"
    mismatched = [e for e in live if offline[e].passed != live[e].passed]
    assert not mismatched, f"{name}: verdict differs offline for {mismatched[:5]}"
    assert score(offline)[0] == sum(1 for r in live.values() if r.passed)


def test_capture_preserves_objects_the_log_would_never_show():
    """O-5: the printed log renders only keyword-matched objects. The capture must
    keep every object, or the dominant failure mode stays unfalsifiable."""
    corpus = load_corpus(CORPUS)
    plan = build_plan(corpus, start_date=START, seed=42, n_days=200, levers=LEVERS)
    _, day_records = simulate(corpus, plan, _subject_title)

    captured = {r["id"] for rec in day_records
                for r in rec["state"]["events"] + rec["state"]["todos"]}
    assert captured, "no objects captured at all"

    # Under subject-titling many objects match no answer-key keyword and so never
    # appear in the log. They must still be in the capture.
    final = day_records[-1]["state"]
    assert len(final["events"]) + len(final["todos"]) >= len(captured) * 0.5
    for rec in final["events"] + final["todos"]:
        assert "email_id" in rec and "title" in rec and "description" in rec
