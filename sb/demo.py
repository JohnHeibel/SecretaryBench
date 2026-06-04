"""
sb.demo — run the deterministic core on the pilot corpus and print a readable
trace: the serve plan, each rendered email, and the oracle's grade.

    python -m sb.demo            # default seed 42
    python -m sb.demo 7          # a different seed -> different serve order
"""
from __future__ import annotations

import sys
from datetime import date

from sb import resolver
from sb.engine import Store, run
from sb.oracle import oracle_model
from sb.resolver import Context
from sb.schema import load_corpus
from sb.scheduler import build_plan

START = date(2026, 6, 1)


def main(seed: int = 42) -> None:
    corpus = load_corpus("corpus")
    plan = build_plan(corpus, start_date=START, seed=seed, n_days=60)

    print(f"\n=== SERVE PLAN  (seed={seed}, start={START}) ===")
    for day, batch in enumerate(plan.per_day):
        if not batch:
            continue
        d = (START.__class__.fromordinal(START.toordinal() + day)).strftime("%a %b %d")
        print(f"  day {day:>2} ({d}): {', '.join(batch)}")
    for name, val in plan.anchors.items():
        print(f"  anchor @{name} = {resolver.human(val)}")
    for eid, dl in plan.deadlines.items():
        print(f"  serve-by deadline: {eid} <= {dl}")

    store = Store(corpus)
    result = run(corpus, plan, oracle_model, store=store)

    print("\n=== SERVE & GRADE (oracle model) ===")
    for day, batch in enumerate(plan.per_day):
        for eid in batch:
            email = corpus.emails[eid]
            ctx = Context(plan.serve_date[eid], plan.anchors)
            rendered = resolver.render_body(email.body, ctx).text
            mark = "PASS" if result.results[eid].passed else "FAIL"
            print(f"  [{mark}] {eid}  ({plan.serve_date[eid]})")
            print(f"         {rendered[:96]}{'...' if len(rendered) > 96 else ''}")

    print(f"\n=== SCORE: {result.passed}/{result.total} = {result.score():.0%} "
          f"(oracle must be 100% on a valid corpus) ===\n")


if __name__ == "__main__":
    main(int(sys.argv[1]) if len(sys.argv) > 1 else 42)
