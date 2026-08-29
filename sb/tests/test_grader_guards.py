"""
The grader's adversarial guard set (docs/grader-contract.md), as tests.

Each world is a synthetic agent that schedules perfectly from the answer key and
varies exactly one thing the grader must, or must not, care about. A grader that
passes these cannot fail a flawless assistant, cannot be bought with volume, cannot
be gamed by a model that never reads a date or that stamps its duplicates with a
sibling's email id, and gives the same verdict on both grading paths and in any
store order. The `real` pin re-grades the certified capture so no grader change
moves the benchmark's score silently.

Everything runs offline, against the corpus and serve plan the capture records.
Numbers marked "recorded" are pinned on purpose: if you changed the grader, move
them deliberately and say why in docs/grader-contract.md; if you changed the
corpus, they move for that reason and the `real` pin skips itself.
"""
import glob
import json
import random
from datetime import date, timedelta
from pathlib import Path

import pytest

from sb.engine import Store, run as engine_run
from sb.grader import grade_email, keywords_of, overlap
from sb.grader import Obj
from sb.live.runner import _corpus_hash, _grade_day, _node_state, _turn_delta
from sb.oracle import _as_dt, _target, oracle_model
from sb.resolver import Context
from sb.schema import load_corpus
from sb.scheduler import Levers, build_plan

CAPTURE = "captures/baseline-sonnet-4-5"
MANIFEST = json.loads(Path(CAPTURE, "manifest.json").read_text())
CORPUS = load_corpus(MANIFEST["corpus_dir"])
_LV = MANIFEST["levers"]
PLAN = build_plan(CORPUS, start_date=date.fromisoformat(MANIFEST["start"]), seed=MANIFEST["seed"],
                  n_days=MANIFEST["n_days"],
                  levers=Levers(_LV["daily_min"], _LV["daily_max"], _LV["urgency_horizon"]))
EMAILS = CORPUS.emails
TOTAL = len(EMAILS)

# Register V-3: what a model that never calls a tool scores. No-op emails pass by
# construction, and so do emails whose every op is a cancel (nothing to cancel).
FLOOR = sum(1 for e in EMAILS.values()
            if not e.answer.ops or all(op.verb == "cancel" for op in e.answer.ops))
MOVE_EMAILS = sum(1 for e in EMAILS.values() if any(op.verb == "move" for op in e.answer.ops))
CANCEL_EMAILS = sum(1 for e in EMAILS.values()
                    if e.answer.ops and all(op.verb == "cancel" for op in e.answer.ops))
# The launderer duplicates on creates only, so its move-only emails may pass.
MOVE_ONLY_EMAILS = sum(1 for e in EMAILS.values()
                       if any(op.verb == "move" for op in e.answer.ops)
                       and not any(op.verb == "create" for op in e.answer.ops))


def _retitled_copy_is_caught(op) -> bool:
    """A stale copy that keeps MORE than half the obligation's words is caught; the
    adversary drops the last word of the title (or appends one to a one-word name)."""
    words = op.name.replace("_", " ").split()
    title = " ".join(words[:-1]) if len(words) > 1 else op.name + " x"
    from datetime import datetime
    return overlap(Obj(kind="event", title=title, when=datetime(2026, 1, 1), email_id=""), keywords_of(op)) > 0.5


RETITLE_CAUGHT = sum(1 for e in EMAILS.values()
                     if any(op.verb == "move" and _retitled_copy_is_caught(op) for op in e.answer.ops))

# Recorded on corpus sha 03e0d963b9866d8f (the 1c capture), grader contract iteration 4.
RECORDED = {"real": 114, "oracle_subject": 137, "oracle_inflect": 165, "wrongkind": 67, "bothkinds": 166,
            "launder_cross": TOTAL}   # see test_launder_across_nodes_is_not_closed_here

_NODE_EMAILS: dict[str, list[str]] = {}
for _e in EMAILS.values():
    _NODE_EMAILS.setdefault(_e.node, []).append(_e.id)


def synth(policy="name", *, act=True, dup=0, shotgun=0, wrongkind=False, dupmove=False,
          retitle_stale=False, launder=None, launder_creates_only=False, skip_cancel=False,
          both_kinds=False):
    """A scheduling-perfect agent with a pluggable title policy, producing day
    records in the capture's shape plus a snapshot after each email (for the
    engine-path score). It remembers which records it made for each obligation,
    so a move or cancel acts on the right object however it was titled -- as a
    real agent does with the ids the store returns.

    dup / shotgun / wrongkind / dupmove / retitle_stale / launder / skip_cancel /
    both_kinds are the adversaries; see docs/grader-contract.md for each.
    """
    events, todos, days, n = [], [], [], 0
    owned: dict[tuple, list[str]] = {}
    for day_no, batch in enumerate([b for b in PLAN.per_day if b], 1):
        before = {r["id"] for r in events + todos}
        per_email = []
        for eid in (batch if act else []):
            em = EMAILS[eid]
            ctx = Context(serve=PLAN.serve_date[eid], anchors=PLAN.anchors)
            before_email = {r["id"] for r in events + todos}
            for op in em.answer.ops:
                base = op.name.replace("_", " ")
                title = {"subject": em.subject,
                         "inflect": " ".join(w + "s" for w in base.split())}.get(policy, base)
                if op.verb == "cancel" and skip_cancel:
                    continue
                if op.verb == "cancel" or (op.verb == "move" and not dupmove and not retitle_stale):
                    gone = set(owned.pop((em.node, op.name), []))
                    events[:] = [r for r in events if r["id"] not in gone]
                    todos[:] = [r for r in todos if r["id"] not in gone]
                    if op.verb == "cancel":
                        continue
                if op.verb == "move" and retitle_stale:
                    for r in events + todos:
                        if r["id"] in set(owned.pop((em.node, op.name), [])):
                            w = r["title"].split()
                            r["title"] = " ".join(w[:-1]) if len(w) > 1 else r["title"] + " x"
                kinds = ["event", "todo"] if both_kinds else [op.kind]
                if wrongkind:
                    kinds = ["todo" if op.kind == "event" else "event"]
                copies = 0 if (launder_creates_only and op.verb == "move") else dup
                made = []
                for kind in kinds:
                    when0 = _as_dt(_target(op.on, ctx), 9 if kind == "event" else 17)
                    for d in (range(shotgun) if shotgun else [0]):
                        when = _as_dt(PLAN.serve_date[eid], 9) + timedelta(days=d) if shotgun else when0
                        for c in range(1 + copies):
                            n += 1
                            stamp = eid
                            if c > 0 and launder == "any":
                                sibs = [x for x in _NODE_EMAILS[em.node] if x != eid]
                                stamp = sibs[0] if sibs else eid
                            elif c > 0 and launder == "cross":
                                stamp = next(x for x in EMAILS if EMAILS[x].node != em.node)
                            elif c > 0 and launder == "past":
                                sibs = [x for x in _NODE_EMAILS[em.node]
                                        if PLAN.serve_date[x] < PLAN.serve_date[eid]]
                                stamp = max(sibs, key=lambda x: PLAN.serve_date[x]) if sibs else eid
                            rec = {"id": f"x_{n}", "email_id": stamp, "title": title, "description": ""}
                            if kind == "event":
                                rec["start"] = when.isoformat()
                                rec["end"] = (when + timedelta(hours=1)).isoformat()
                                events.append(rec)
                            else:
                                rec["due_date"] = when.isoformat()
                                todos.append(rec)
                            made.append(rec["id"])
                owned.setdefault((em.node, op.name), []).extend(made)
            per_email.append({"eid": eid, "state": {"events": list(events), "todos": list(todos)},
                              "new": sorted({r["id"] for r in events + todos} - before_email)})
        days.append({"batch": list(batch), "state": {"events": list(events), "todos": list(todos)},
                     "day_new": sorted({r["id"] for r in events + todos} - before),
                     "per_email": per_email})
    return days


def score_runner(days) -> int:
    """The live runner / sb.regrade path: one day-end state, graded by _grade_day."""
    return sum(bool(r.passed)
               for rec in days
               for r in _grade_day(CORPUS, PLAN, rec["batch"], rec["state"], set(rec["day_new"])).values())


def score_engine(days) -> int:
    """The sb.engine path: each email graded right after its own turn, against the
    state as it stood then, with its whole delta as its turn (and so its day)."""
    passed = 0
    for rec in days:
        for pe in rec["per_email"]:
            em = EMAILS[pe["eid"]]
            ctx = Context(serve=PLAN.serve_date[pe["eid"]], anchors=PLAN.anchors)
            turn = _turn_delta(CORPUS, pe["state"], set(pe["new"]))
            state = _node_state(CORPUS, pe["state"], em.node, set(pe["new"]))
            passed += bool(grade_email(em.answer, ctx, state, turn).passed)
    return passed


def shuffled(days, seed):
    rnd = random.Random(seed)
    out = []
    for rec in days:
        st = {"events": list(rec["state"]["events"]), "todos": list(rec["state"]["todos"])}
        rnd.shuffle(st["events"])
        rnd.shuffle(st["todos"])
        out.append({**rec, "state": st})
    return out


# --- a flawless assistant must get full marks -------------------------------

def test_oracle_through_the_engine_scores_full_marks():
    """sb.scale's mandatory gate, on the capture's plan (register G-4)."""
    res = engine_run(CORPUS, PLAN, oracle_model, store=Store(CORPUS))
    assert res.passed == res.total == TOTAL, [e for e, r in res.results.items() if not r.passed]


def test_perfect_agent_titling_by_obligation_name_scores_full_marks():
    assert score_runner(synth("name")) == TOTAL


# --- volume must not buy the score --------------------------------------------

def test_null_model_scores_exactly_the_floor():
    assert score_runner(synth(act=False)) == FLOOR


@pytest.mark.parametrize("world,kw", [
    ("dup5", dict(dup=5)),                                   # right answer, then five copies of every object
    ("shot7", dict(shotgun=7)),                              # never reads a date: one object per day for a week
    ("shot90", dict(shotgun=90)),                            # ... for ninety days
    ("launder_all", dict(dup=5, launder="any")),             # five copies, stamped with a node sibling's id
    ("launder_past", dict(dup=5, launder="past")),           # ... stamped with an already-graded sibling's id
])
def test_volume_never_beats_the_null_floor(world, kw):
    assert score_runner(synth("name", **kw)) <= FLOOR, world


def test_laundered_duplicates_on_creates_keep_only_the_move_only_emails():
    """Copies stamped with a sibling's id, on creates only: every email with a
    create fails; the move-only emails, where this agent leaves nothing, may pass."""
    assert score_runner(synth("name", dup=5, launder="any", launder_creates_only=True)) <= FLOOR + MOVE_ONLY_EMAILS


def test_launder_across_nodes_is_not_closed_here():
    """Copies stamped with ANOTHER node's email id never reach this node's pool:
    sb.live.runner._node_state drops them by attribution before the grader sees
    them, so this agent scores full marks with five times the calendar. That is
    register A-5, which the identity contract does not close (phase 3). Pinned so
    the day it changes, someone changed attribution on purpose."""
    assert score_runner(synth("name", dup=5, launder="cross")) == RECORDED["launder_cross"]


def test_every_double_booked_move_fails_and_nothing_else_does():
    """A model that 'moves' by creating a copy and leaving the old object loses
    exactly the emails that contain a move op (contract rule 4)."""
    assert score_runner(synth("name", dupmove=True)) == TOTAL - MOVE_EMAILS


def test_a_retitled_stale_copy_is_still_caught_while_it_keeps_most_of_the_name():
    assert score_runner(synth("name", retitle_stale=True)) <= TOTAL - RETITLE_CAUGHT


def test_never_deleting_anything_fails_every_cancel_email():
    assert score_runner(synth("name", skip_cancel=True)) == TOTAL - CANCEL_EMAILS


# --- pinned worlds: move these deliberately -----------------------------------

@pytest.mark.parametrize("world,kw", [
    ("oracle_subject", dict(policy="subject")),    # realistic naming: titles = email subject
    ("oracle_inflect", dict(policy="inflect")),    # every title word pluralised
    ("wrongkind", dict(wrongkind=True)),           # right title and day, wrong kind
    ("bothkinds", dict(both_kinds=True)),          # an event AND a to-do for every obligation
])
def test_recorded_worlds_do_not_move_silently(world, kw):
    got = score_runner(synth(**kw))
    assert got == RECORDED[world], (
        f"{world}: {got}, recorded {RECORDED[world]}. If the grader changed, re-baseline in "
        f"docs/grader-contract.md; if the corpus changed, that is the cause.")


def test_certified_capture_regrades_to_the_recorded_score():
    if _corpus_hash(MANIFEST["corpus_dir"]) != MANIFEST["corpus_hash"]:
        pytest.skip("corpus differs from the one the capture was served from; "
                    "the capture cannot be re-graded against it")
    days = [json.loads(Path(p).read_text()) for p in sorted(glob.glob(f"{CAPTURE}/days/*.json"))]
    got = score_runner(days)
    assert got == RECORDED["real"], (
        f"real capture re-grades to {got}, recorded {RECORDED['real']}: a grader change moved "
        f"the benchmark's score. Re-baseline deliberately in docs/grader-contract.md.")


# --- the same state must grade the same on both paths and in any order --------

@pytest.mark.parametrize("world,kw", [
    ("oracle_name", dict()),
    ("oracle_subject", dict(policy="subject")),
    ("dupmove", dict(dupmove=True)),
    ("launder", dict(dup=5, launder="any", launder_creates_only=True)),
    ("launder_past", dict(dup=5, launder="past")),
    ("nocancel", dict(skip_cancel=True)),
])
def test_engine_and_runner_paths_agree(world, kw):
    """sb.engine grades each email from its own turn; the live runner and sb.regrade
    grade a day-end state split back to each email by its stamped email_id.
    Iteration 3 of the contract scored one adversary 64 on the first path and 167
    on the second (VERIFY-phase2-iter3 §2)."""
    days = synth(**kw)
    assert score_engine(days) == score_runner(days), world


@pytest.mark.parametrize("world,kw", [
    ("oracle_name", dict()),
    ("oracle_subject", dict(policy="subject")),
    ("dupmove", dict(dupmove=True)),
])
def test_verdicts_are_invariant_to_store_order(world, kw):
    """The store lists objects in insertion order and delete-then-recreate changes
    it; the grade must not (sb/grader.py's stated contract; VERIFY-phase2-iter3 §3)."""
    days = synth(**kw)
    canonical = score_runner(days)
    assert {score_runner(shuffled(days, s)) for s in range(4)} == {canonical}, world
