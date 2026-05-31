"""Shared static facts: one named scalar drives both the rendered prose and the
answer keys of dependent emails, so a policy change propagates everywhere."""
from __future__ import annotations

import json
from datetime import date, datetime

from sb.grader import NodeState, Obj, grade_email
from sb.resolver import Context, render_body
from sb.schema import load_corpus


def _write_corpus(tmp_path, meeting_len: str):
    src = load_corpus("corpus")  # start from the real corpus, then rewrite the fact value
    dest = tmp_path / "corpus"
    (dest / "nodes").mkdir(parents=True)
    for nid in src.nodes:
        raw = json.loads(open(f"corpus/nodes/{nid}.json").read())
        if nid == "hr-policy":
            raw["emails"][0]["body"] = raw["emails"][0]["body"].replace("= 90m}", f"= {meeting_len}}}")
        (dest / "nodes" / f"{nid}.json").write_text(json.dumps(raw))
    return load_corpus(dest)


def test_fact_renders_into_prose(tmp_path):
    corpus = _write_corpus(tmp_path, "45m")
    memo = corpus.emails["hr-policy.memo"]
    text = render_body(memo.body, Context(date(2026, 6, 1), {}, corpus.facts)).text
    assert "45-minute blocks" in text
    assert corpus.facts["client_meeting_len"] == "45m"


def test_policy_change_propagates_to_answer_key(tmp_path):
    # With the policy at 45m, a 90-minute acme event must now FAIL; a 45-minute one passes.
    corpus = _write_corpus(tmp_path, "45m")
    acme = corpus.emails["acme-client.request"]
    ctx = Context(date(2026, 6, 1), corpus.emails["acme-client.request"].emits, corpus.facts)
    # build the expected start from the answer predicate
    from sb.resolver import resolve
    start = resolve(acme.answer.expect[0].start["eq"], ctx)
    assert isinstance(start, datetime)

    def grade(minutes: int):
        ev = Obj(kind="event", title="Acme renewal", when=start, email_id=acme.id,
                 end=start.replace() + __import__("datetime").timedelta(minutes=minutes))
        state = NodeState(events=[ev])
        from sb.grader import TurnDelta
        return grade_email(acme.answer, ctx, state, TurnDelta(events=[ev]))

    assert grade(45).passed, "45-minute event should pass under a 45m policy"
    assert not grade(90).passed, "90-minute event should fail once policy says 45m"


def test_default_corpus_fact_is_90m():
    corpus = load_corpus("corpus")
    assert corpus.facts["client_meeting_len"] == "90m"
    assert corpus.fact_map["client_meeting_len"] == "hr-policy.memo"
