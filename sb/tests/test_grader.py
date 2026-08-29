"""Focused grader/oracle edge cases."""
from datetime import date, datetime

import pytest

from sb.grader import NodeState, Obj, TurnDelta, grade_email, keywords_of
from sb.oracle import _target
from sb.resolver import Context, Interval
from sb.schema import Answer, Op


def test_by_predicate_rejects_due_date_before_email_arrives():
    ctx = Context(serve=date(2026, 6, 10))
    answer = Answer(ops=[Op(verb="create", name="filing", kind="todo", match=["filing"], on={"by": "serve+5d"})])
    todo = Obj(kind="todo", title="filing", when=datetime(2026, 6, 9, 17), email_id="n.a")

    # the to-do was created this turn (a create claims only what was made today)
    result = grade_email(answer, ctx, NodeState(todos=[todo]), TurnDelta(todos=[todo]))

    assert not result.passed
    assert result.details[0]["reason"] == "on the wrong day"


def test_by_predicate_accepts_due_date_between_serve_and_deadline():
    ctx = Context(serve=date(2026, 6, 10))
    answer = Answer(ops=[Op(verb="create", name="filing", kind="todo", match=["filing"], on={"by": "serve+5d"})])
    todo = Obj(kind="todo", title="filing", when=datetime(2026, 6, 12, 17), email_id="n.a")

    assert grade_email(answer, ctx, NodeState(todos=[todo]), TurnDelta(todos=[todo])).passed


def test_oracle_target_avoids_not_in_blackout():
    ctx = Context(serve=date(2026, 6, 1), anchors={"blackout": Interval(date(2026, 6, 8), date(2026, 6, 10))})

    assert _target({"in": "week_of:(serve+1w)", "not_in": "@blackout"}, ctx) == date(2026, 6, 11)


def test_oracle_target_errors_when_window_fully_blocked():
    ctx = Context(serve=date(2026, 6, 1), anchors={"blackout": Interval(date(2026, 6, 8), date(2026, 6, 14))})

    with pytest.raises(ValueError, match="fully blocked"):
        _target({"in": "week_of:(serve+1w)", "not_in": "@blackout"}, ctx)



# --- the identity contract (docs/grader-contract.md; register G-10) -----------

def _ctx():
    return Context(serve=date(2026, 6, 10))


def _ev(title, day, email_id="n.a", description=""):
    return Obj(kind="event", title=title, when=datetime(2026, 6, day, 9), email_id=email_id,
               description=description)


def _td(title, day, email_id="n.a", description=""):
    return Obj(kind="todo", title=title, when=datetime(2026, 6, day, 17), email_id=email_id,
               description=description)


def _create(name, kind="event", on=None, **kw):
    return Op(verb="create", name=name, kind=kind, on=on or {"eq": "serve+2d"}, **kw)


def _grade(ops, events=(), todos=(), fresh=()):
    """Grade one email whose turn (and day) created `fresh`."""
    state = NodeState(events=list(events), todos=list(todos))
    turn = TurnDelta(events=[o for o in fresh if o.kind == "event"], todos=[o for o in fresh if o.kind == "todo"])
    return grade_email(Answer(ops=list(ops)), _ctx(), state, turn)


def test_keywords_are_the_stemmed_content_words_of_the_name():
    assert keywords_of(_create("Schedule the Board Meeting")) == {"board"}
    assert keywords_of(_create("Delayed_Release_Date")) == {"delay", "release"}
    assert keywords_of(_create("Send trophies and notes")) == {"trophy", "note"}
    # a name made only of stop-words falls back to all of its words
    assert keywords_of(_create("event day!")) == {"event", "day"}
    # stop-words are matched on the raw word: an inflected one survives, stemmed
    assert keywords_of(_create("Board Meetings")) == {"board", "meet"}


def test_identity_is_word_level_and_ignores_the_match_field():
    # "sign-off" vs "signoff" fails the shipped substring rule on one hyphen (G-1);
    # here the shared content word finds it and the date decides.
    obj = _ev("Board sign-off meeting", 12)
    res = _grade([_create("board signoff", match=["zzz-never-in-any-title"])], events=[obj], fresh=[obj])
    assert res.passed, res.headline


def test_description_completes_a_match_but_title_evidence_ranks_first():
    obj = _ev("Partnership call", 12, description="With the WHOOP team")
    assert _grade([_create("whoop meeting")], events=[obj], fresh=[obj]).passed
    # two candidates: one names the obligation in its title, one only in its description
    by_title = _ev("FBS presentation", 12)
    by_desc = _ev("Delegation meeting", 12, description="prep for the FBS conference")
    res = _grade([_create("FBS Conference")], events=[by_desc, by_title], fresh=[by_desc, by_title])
    assert "FBS presentation" in res.details[0]["actual"]


def test_distinct_obligations_sharing_vocabulary_are_not_duplicates():
    """G-2: 0 of 57 recorded 'duplicates' were duplicates. Exclusive assignment
    gives each obligation its own object; the more specific match claims first."""
    objs = [_ev("Reveal event on stage", 12), _ev("Reveal rehearsal", 12)]
    res = _grade([_create("Reveal Event"), _create("Reveal Rehearsal")], events=objs, fresh=objs)
    assert res.passed, res.headline
    assert "rehearsal" in res.details[1]["actual"].lower()


def test_a_duplicate_created_today_fails():
    objs = [_ev("Atlas launch dinner", 12), _ev("Atlas launch dinner", 12)]
    res = _grade([_create("Atlas launch dinner")], events=objs, fresh=objs)
    assert not res.passed and "over-created" in res.headline


def test_a_move_that_leaves_the_old_object_fails_but_a_clean_move_passes():
    move = Op(verb="move", name="sync", kind="event", on={"eq": "serve+5d"})
    old = _ev("sync", 4, email_id="n.book")
    new = _ev("sync", 15, email_id="n.shift")
    # recreated and the old one left behind: double-booked
    res = _grade([move], events=[old, new], fresh=[new])
    assert not res.passed and "stale copy" in res.headline
    # ... even if the stale copy was retitled, while it keeps more than half the words
    old2 = _ev("Vendor kickoff", 4, email_id="n.book")                 # 2 of 3 words kept
    move2 = Op(verb="move", name="Vendor kickoff prep", kind="event", on={"eq": "serve+5d"})
    new2 = _ev("Vendor kickoff prep", 15, email_id="n.shift")
    assert not _grade([move2], events=[old2, new2], fresh=[new2]).passed
    # updated in place (nothing new today): fine
    assert _grade([move], events=[_ev("sync", 15, email_id="n.book")]).passed
    # deleted and recreated: fine
    assert _grade([move], events=[new], fresh=[new]).passed


def test_a_create_claims_only_what_was_made_today():
    """An inherited look-alike is an earlier obligation's answer (G-2), never a
    fresh create's; a create that made nothing is 'not found', not 'wrong day'."""
    old = _ev("Retreat Company Meeting Call", 3, email_id="n.earlier")
    new = _ev("Company Retreat", 12, email_id="n.now")
    assert _grade([_create("Company Retreat")], events=[old, new], fresh=[new]).passed
    res = _grade([_create("Company Retreat")], events=[old])
    assert not res.passed and res.headline.startswith("no event")


def test_cancel_passes_when_only_a_sibling_sharing_a_word_remains():
    """G-7: a sibling the model was told to keep must not fail the cancel."""
    cancel = Op(verb="cancel", name="Design Lead Stage Slot", kind="event")
    kept = _ev("Design walk-through", 12, description="Product design review")
    assert _grade([cancel], events=[kept]).passed


def test_cancel_fails_on_a_same_kind_survivor_carrying_every_word():
    cancel = Op(verb="cancel", name="Design Lead Stage Slot", kind="event")
    survivor = _ev("Stage slot for the design lead", 12)
    res = _grade([cancel], events=[survivor])
    assert not res.passed and "still on the calendar" in res.headline
    # a survivor of the other kind is not the cancel's business (the shipped rule; G-7)
    assert _grade([cancel], todos=[_td("Stage slot for the design lead", 12)]).passed


def test_a_sibling_op_cannot_shield_a_cancel_by_a_weaker_claim():
    """The object that carries all four words of the cancelled obligation goes to
    the cancel, not to a sibling move that shares one word with it."""
    cancel = Op(verb="cancel", name="Design Lead Stage Slot", kind="event")
    move = Op(verb="move", name="Reveal Event", kind="event", on={"eq": "serve+3d"})
    slot = _ev("Design lead stage slot at the reveal", 12)
    reveal = _ev("Reveal event", 13)
    res = _grade([move, cancel], events=[slot, reveal])
    assert res.details[0]["passed"], res.details[0]                # the move took the reveal
    assert not res.details[1]["passed"], res.details[1]            # the cancel found its survivor


def test_wrong_kind_is_reported_not_claimed():
    wrong = _td("Podcast taping", 12)
    res = _grade([_create("podcast taping", kind="event")], todos=[wrong], fresh=[wrong])
    assert not res.passed and res.headline.startswith("wrong kind")
    # a right-kind object, however weakly titled, is the one that counts
    right = _ev("Taping", 12)
    res = _grade([_create("podcast taping", kind="event")], events=[right], todos=[wrong], fresh=[right, wrong])
    assert res.passed, res.headline
    # a sibling's claimed object is never reported as this op's wrong-kind attempt
    ops = [_create("Create a list for the athlete visit", kind="todo"),
           _create("Contact people added to the list", kind="event")]
    both = _td("Create guest list for athlete visit and contact invitees", 12)
    res = _grade(ops, todos=[both], fresh=[both])
    assert res.details[1]["reason"].startswith("no event")


def test_verdicts_do_not_depend_on_pool_order():
    ops = [_create("Reveal Event"), _create("Reveal Rehearsal"), _create("Press briefing")]
    objs = [_ev("Reveal rehearsal with striker", 12), _ev("Press briefing", 12), _ev("Cleat reveal event", 12)]
    a = _grade(ops, events=objs, fresh=objs)
    b = _grade(ops, events=objs[::-1], fresh=objs)
    assert [d["reason"] for d in a.details] == [d["reason"] for d in b.details] == ["matched"] * 3
    # ... including when two objects differ only in their description (VERIFY-phase2-iter4 §5.2)
    ops = [_create("Team Offsite"), _create("Budget")]
    plain = _ev("Team Offsite", 12)
    budget = _ev("Team Offsite", 12, description="budget review")
    a = _grade(ops, events=[plain, budget], fresh=[plain, budget])
    b = _grade(ops, events=[budget, plain], fresh=[plain, budget])
    assert [d["passed"] for d in a.details] == [d["passed"] for d in b.details] == [True, True]


def test_a_perfect_tie_goes_to_the_work_not_the_cancel():
    """Same keyword set, one fresh object: the create gets it, the cancel passes —
    whichever email the object is stamped to (VERIFY-phase2-iter4 §5.3)."""
    ops = [Op(verb="cancel", name="Vendor Onsite", kind="event"), _create("Vendor Onsite")]
    for stamp in ("n.create", "n.cancel", "n.other"):
        obj = _ev("Vendor onsite", 12, email_id=stamp)
        res = _grade(ops, events=[obj], fresh=[obj])
        assert res.passed, (stamp, res.headline)


def test_the_more_specific_obligation_claims_first_whatever_the_op_order():
    """Two obligations, one nested in the other, competing for the same objects."""
    rehearsal = _ev("Reveal rehearsal", 12)
    event = _ev("Reveal event on stage", 12)
    for ops in ([_create("Reveal Event"), _create("Reveal Rehearsal")],
                [_create("Reveal Rehearsal"), _create("Reveal Event")]):
        res = _grade(ops, events=[rehearsal, event], fresh=[rehearsal, event])
        assert res.passed, res.headline
        for d in res.details:
            assert ("rehearsal" in d["label"].lower()) == ("rehearsal" in d["actual"].lower()), d


def test_duplicates_are_judged_against_the_whole_day():
    """A copy stamped with a sibling's id is still a duplicate (register A-5); a
    sibling's own object, which does not match this obligation as well, is not."""
    mine = _ev("Vendor sync", 12, email_id="n.a")
    theirs = _ev("Budget memo review", 12, email_id="n.b")
    laundered = _ev("Vendor sync", 12, email_id="n.b")
    state = NodeState(events=[mine, theirs, laundered])
    today = TurnDelta(events=[mine, theirs, laundered])
    res = grade_email(Answer(ops=[_create("Vendor Sync")]), _ctx(), state, TurnDelta(events=[mine]), today)
    assert not res.passed and "over-created" in res.headline
    state = NodeState(events=[mine, theirs])
    today = TurnDelta(events=[mine, theirs])
    assert grade_email(Answer(ops=[_create("Vendor Sync")]), _ctx(), state, TurnDelta(events=[mine]), today).passed
    assert grade_email(Answer(ops=[_create("Budget Memo")]), _ctx(), state, TurnDelta(events=[theirs]), today).passed
