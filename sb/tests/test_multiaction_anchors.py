"""
Validates the webapp authoring feature where ONE email's body publishes several
timed `{!name=...}` anchors and that SAME email's answer creates one event per
anchor (reusing `@name`). Proves the real reference oracle grades it 1.0, and
exercises the failure-reporting path the webapp's /api/oracle handler builds.

Also probes (does NOT fix) a latent scheduler deadlock: an email that self-emits
the anchors its own answer references must NOT carry a `date` edge, or build_plan
strands it as InfeasibleSchedule (see test_step4_self_emit_with_date_edge_deadlocks).

Pattern mirrors test_e2e.py: build_corpus -> build_plan(start 2026-06-01, seed 42)
-> engine.run(corpus, plan, oracle_model); a passed Store lets us read placements.
"""
from datetime import date, datetime, timedelta

import pytest

from sb import engine, resolver, schema
from sb.oracle import oracle_model
from sb.resolver import Context, TimeInterval
from sb.scheduler import InfeasibleSchedule, build_plan

START = date(2026, 6, 1)


def _plan(corpus, n_days=60):
    return build_plan(corpus, start_date=START, seed=42, n_days=n_days)


# --- Step 1: self-contained multi-action, same-email timed anchors ----------

def _step1_node():
    return {
        "id": "avail",
        "cast": {"YOU": "you", "BOSS": "Boss"},
        "emails": [{
            "id": "avail.req", "from": "BOSS", "to": "YOU", "subject": "availabilities",
            "body": ("Serena can meet {!sw_time = next:MON @10:00-11:00}. "
                     "Michael can do {!mp_time = next:TUE @09:00-10:00}. "
                     "Shaq might make {!so_time = next:MON @13:00-14:00}."),
            "depends_on": [],
            "answer": {"ops": [
                {"create": "serena_meeting", "kind": "event", "match": ["serena"], "on": {"eq": "@sw_time"}},
                {"create": "phelps_meeting", "kind": "event", "match": ["michael"], "on": {"eq": "@mp_time"}},
                {"create": "shaq_meeting", "kind": "event", "match": ["shaq"], "on": {"eq": "@so_time"}},
            ]},
        }],
    }


def test_step1_same_email_timed_anchors_oracle_perfect():
    corpus = schema.build_corpus([_step1_node()])            # lint clean (no raise)
    plan = _plan(corpus)
    store = engine.Store(corpus)
    res = engine.run(corpus, plan, oracle_model, store=store)

    assert res.total == 1
    assert res.score() == 1.0

    # Served on Mon Jun 1: next:MON is strictly AFTER -> Jun 8; next:TUE -> Jun 2.
    sd = plan.serve_date["avail.req"]
    assert sd == date(2026, 6, 1)
    ctx = Context(serve=sd, anchors=plan.anchors)
    assert resolver.resolve("@sw_time", ctx) == TimeInterval(datetime(2026, 6, 8, 10), datetime(2026, 6, 8, 11))
    assert resolver.resolve("@mp_time", ctx) == TimeInterval(datetime(2026, 6, 2, 9), datetime(2026, 6, 2, 10))
    assert resolver.resolve("@so_time", ctx) == TimeInterval(datetime(2026, 6, 8, 13), datetime(2026, 6, 8, 14))

    # The three events landed at EXACT start/end (timed, minute-granular).
    events = {e.title: (e.when, e.end) for e in store.node_state("avail").events}
    assert events["serena"] == (datetime(2026, 6, 8, 10), datetime(2026, 6, 8, 11))
    assert events["michael"] == (datetime(2026, 6, 2, 9), datetime(2026, 6, 2, 10))
    assert events["shaq"] == (datetime(2026, 6, 8, 13), datetime(2026, 6, 8, 14))


# --- Step 2: cross-email anchor + same-email reuse together -----------------

def _step2_node():
    return {
        "id": "screenshot",
        "cast": {"YOU": "you", "BOSS": "Boss"},
        "emails": [
            {"id": "ss.move", "from": "BOSS", "to": "YOU", "subject": "reschedule",
             "body": "Let's move things to {!rescheduled_date = next:MON}.",
             "depends_on": [], "answer": {"ops": []}},
            {"id": "ss.avail", "from": "BOSS", "to": "YOU", "subject": "availabilities",
             "body": ("Serena {!sw_time = @rescheduled_date+2d @10:00-11:00}. "
                      "Michael {!mp_time = @rescheduled_date+3d @09:00-10:00}. "
                      "Shaq {!so_time = @rescheduled_date+2d @13:00-14:00}."),
             # MUST be static, not date: ss.avail self-emits the anchors its own answer
             # references, so a date edge would deadlock the scheduler (see step 4).
             "depends_on": [{"email": "ss.move", "type": "static"}],
             "answer": {"ops": [
                 {"create": "serena_meeting", "kind": "event", "match": ["serena"], "on": {"eq": "@sw_time"}},
                 {"create": "phelps_meeting", "kind": "event", "match": ["michael"], "on": {"eq": "@mp_time"}},
                 {"create": "shaq_meeting", "kind": "event", "match": ["shaq"], "on": {"eq": "@so_time"}},
             ]}},
        ],
    }


def test_step2_cross_email_anchor_plus_self_emit_oracle_perfect():
    corpus = schema.build_corpus([_step2_node()])
    plan = _plan(corpus)
    store = engine.Store(corpus)
    res = engine.run(corpus, plan, oracle_model, store=store)

    assert res.total == 2
    assert res.score() == 1.0

    # ss.move served Jun 1 -> rescheduled_date = next:MON = Jun 8. +2d=Jun 10, +3d=Jun 11.
    assert plan.anchors["rescheduled_date"] == date(2026, 6, 8)
    events = {e.title: (e.when, e.end) for e in store.node_state("screenshot").events}
    assert events["serena"] == (datetime(2026, 6, 10, 10), datetime(2026, 6, 10, 11))
    assert events["michael"] == (datetime(2026, 6, 11, 9), datetime(2026, 6, 11, 10))
    assert events["shaq"] == (datetime(2026, 6, 10, 13), datetime(2026, 6, 10, 14))


# --- Step 3: failure reporting (the webapp/api/oracle.py `failures` shape) ---

def _oracle_failures(corpus, res):
    """Reconstruct exactly the list webapp/api/oracle.py `_oracle` builds from a run:
    one {id, node, reason} per failing email, where reason = EmailResult.headline."""
    return [
        {"id": eid, "node": corpus.emails[eid].node if eid in corpus.emails else "", "reason": r.headline}
        for eid, r in res.results.items() if not r.passed
    ]


def test_step3_oracle_happy_path_has_zero_failures():
    # The valid corpus -> the API would return ok:true, failures:[].
    corpus = schema.build_corpus([_step1_node()])
    res = engine.run(corpus, _plan(corpus), oracle_model)
    assert res.score() == 1.0
    assert _oracle_failures(corpus, res) == []


def test_step3_imperfect_model_failures_are_well_formed():
    """An imperfect model that double-books every event must fail, and the failing
    EmailResult must carry a non-empty .headline (the 'reason' the API surfaces)."""
    from sb.oracle import _as_dt, _target

    def double_booker(email, rendered, ctx, store):
        for op in email.answer.ops:
            title = " ".join(op.match) if op.match else op.name
            when = _as_dt(_target(op.on, ctx), 9)
            store.create_event(email.id, title, when)
            store.create_event(email.id, title, when)        # duplicate -> count != 1

    corpus = schema.build_corpus([_step1_node()])
    res = engine.run(corpus, _plan(corpus), double_booker)

    failing = {eid: r for eid, r in res.results.items() if not r.passed}
    assert failing, "imperfect model should fail at least one email"
    for r in failing.values():
        assert isinstance(r.headline, str) and r.headline    # non-empty reason
    assert res.results["avail.req"].headline == "found 2 matching, expected exactly 1 (duplicate / double-booked)"

    failures = _oracle_failures(corpus, res)
    assert failures and all(set(f) == {"id", "node", "reason"} for f in failures)
    for f in failures:
        assert f["id"] in corpus.emails
        assert f["node"] == "avail"
        assert f["reason"]


def test_step3_wrong_clock_time_surfaces_timed_reason():
    """A single create on the right DAY but the wrong clock time yields the timed
    headline ('at the wrong time'), not the day-level 'on the wrong day'."""
    def wrong_hour(email, rendered, ctx, store):
        for op in email.answer.ops:
            title = " ".join(op.match) if op.match else op.name
            v = resolver.resolve(op.on["eq"], ctx)           # a TimeInterval
            bad = v.start + timedelta(hours=5)
            store.create_event(email.id, title, bad, bad + timedelta(hours=1))

    corpus = schema.build_corpus([_step1_node()])
    res = engine.run(corpus, _plan(corpus), wrong_hour)
    assert not res.results["avail.req"].passed
    assert res.results["avail.req"].headline == "at the wrong time"


# --- Step 4: latent scheduler deadlock (REPORT ONLY — do not fix) -----------

def _step4_node():
    """Step-2 node plus a 4th create op reusing the ANCESTOR anchor @rescheduled_date
    directly. That makes ss.avail a true needle to ss.move, so a `date` edge is the
    'correct' wiring per the needle rule — but ss.avail also self-emits sw/mp/so which
    its own answer references."""
    node = _step2_node()
    node["id"] = "screenshot4"
    node["emails"][0]["id"] = "s4.move"
    b = node["emails"][1]
    b["id"] = "s4.avail"
    b["depends_on"] = [{"email": "s4.move", "type": "date"}]   # date edge (the needle wiring)
    b["answer"]["ops"].append(
        {"create": "prep", "kind": "event", "match": ["prep"], "on": {"eq": "@rescheduled_date"}})
    return node


def test_step4_self_emit_with_date_edge_deadlocks():
    """REPORT: this corpus passes lint (build_corpus does NOT raise) but build_plan
    raises InfeasibleSchedule. Root cause: ss.avail is a `date`-edge email, so
    scheduler.update_deadlines refuses to deadline it until ALL its anchor_refs are
    in the anchor table — but those refs include sw/mp/so, which ss.avail only emits
    WHEN it is served. is_ready never releases it -> stranded forever. A self-emitting
    email must use a `static` edge (as step 2 does). Latent engine gotcha, not fixed here."""
    corpus = schema.build_corpus([_step4_node()])            # lint passes
    # ss.avail picked up a derived `date` edge and references its own self-emitted anchors.
    assert any(e.type == "date" for e in corpus.emails["s4.avail"].depends_on)
    assert {"sw_time", "mp_time", "so_time"} <= corpus.emails["s4.avail"].anchor_refs

    with pytest.raises(InfeasibleSchedule) as exc:
        _plan(corpus)
    assert "s4.avail" in str(exc.value)
