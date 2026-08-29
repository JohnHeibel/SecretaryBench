"""
End-to-end dry run of the deterministic core: load -> plan -> serve -> grade,
with mock models standing in for the harness. Proves the whole pipeline works
and that the grader distinguishes a correct model from a broken one.
"""
from datetime import date
from pathlib import Path

from sb.engine import Store, run
from sb.oracle import _as_dt, _target, oracle_model as perfect_model
from sb.resolver import Context
from sb.schema import Email, load_corpus
from sb.scheduler import build_plan

START = date(2026, 6, 1)
FIXTURE = str(Path(__file__).parent / "fixtures")


def imperfect_model(email: Email, rendered: str, ctx: Context, store: Store) -> None:
    """A plausible-but-wrong model:
      - over-acts on no-action mail (creates a stray todo on every FYI)
      - 'moves' by creating a second object instead of moving the first,
        so any move leaves a duplicate (more than the expected single object)
    """
    ops = email.answer.ops
    if not ops:
        store.create_todo(email.id, "follow up", _as_dt(ctx.serve, 17))   # over-action
        return
    for op in ops:
        title = op.name.replace("_", " ")
        if op.verb == "cancel":
            continue                                       # never cancels -> leaves the object
        when = _as_dt(_target(op.on, ctx), 9 if op.kind == "event" else 17)
        if op.kind == "event":
            store.create_event(email.id, title, when)      # always creates -> double-books on a move
        else:
            store.create_todo(email.id, title, when)


def _run(model):
    corpus = load_corpus(FIXTURE)
    plan = build_plan(corpus, start_date=START, seed=42, n_days=60)
    return corpus, run(corpus, plan, model)


def test_perfect_model_scores_full_marks():
    _corpus, res = _run(perfect_model)
    assert res.total == 6
    failures = {eid for eid, r in res.results.items() if not r.passed}
    assert failures == set(), f"perfect model should pass everything, failed: {failures}"
    assert res.score() == 1.0


def test_imperfect_model_is_caught():
    _corpus, res = _run(imperfect_model)
    # over-action on a no-action FYI, and a reschedule done as a duplicate, must fail.
    assert not res.results["gamma.notice"].passed           # no-action violated
    # Double-booked reschedule. Re-baselined 2026-08-29 with the identity contract
    # (docs/grader-contract.md rule 4): it still fails, now as a stale copy left
    # behind by the move rather than under the old exactly-one count.
    shift = res.results["beta.shift"]
    assert not shift.passed
    assert "stale copy" in shift.headline, shift.headline
    # a plain single create on the right day still passes (day-level grading)
    assert res.results["alpha.brief"].passed
    assert res.results["alpha.review"].passed
    assert 0.0 < res.score() < 1.0


def test_reschedule_via_update_keeps_count_one():
    corpus = load_corpus(FIXTURE)
    plan = build_plan(corpus, start_date=START, seed=42, n_days=60)
    store = Store(corpus)
    run(corpus, plan, perfect_model, store=store)
    beta_events = store.node_state("beta").events
    assert len(beta_events) == 1               # moved, not duplicated
