"""
sb.span — measure retrieval span: for every email whose ANSWER references an
anchor, how far (in emails served and in days) is it from the email that emitted
that anchor? Large span = the fact has likely scrolled out of context by the time
it's needed, so the model must search_inbox to recover it. This is the axis the
benchmark's research claim rests on (see project docs), so we measure it directly.

    python -m sb.span [seed] [n_days]
"""
from __future__ import annotations

import sys
from datetime import date

from sb.schema import Corpus, _email_predicates, _refs_in, load_corpus
from sb.scheduler import build_plan, Plan


def _answer_anchor_refs(corpus: Corpus, email_id: str) -> set[str]:
    refs: set[str] = set()
    for expr in _email_predicates(corpus.emails[email_id].answer):
        refs |= _refs_in(expr)
    return refs


def spans(corpus: Corpus, plan: Plan) -> list[dict]:
    """One record per (referencing email, anchor it needs)."""
    order = [eid for batch in plan.per_day for eid in batch]
    pos = {eid: i for i, eid in enumerate(order)}
    out: list[dict] = []
    for eid in order:
        for name in _answer_anchor_refs(corpus, eid):
            src = corpus.emission_map.get(name)
            if not src or src == eid:
                continue
            out.append({
                "needs": eid, "anchor": name, "from": src,
                "email_span": pos[eid] - pos[src],
                "day_span": (plan.serve_date[eid] - plan.serve_date[src]).days,
            })
    return out


def main(seed: int = 42, n_days: int = 100) -> None:
    corpus = load_corpus("corpus")
    plan = build_plan(corpus, start_date=date(2026, 6, 1), seed=seed, n_days=n_days)
    recs = sorted(spans(corpus, plan), key=lambda r: r["email_span"], reverse=True)
    total = len([e for b in plan.per_day for e in b])
    print(f"\ncorpus: {len(corpus.emails)} emails, {total} served (seed={seed})")
    if not recs:
        print("no cross-email anchor references found.")
        return
    print(f"{'needs':<26} {'anchor':<18} {'from':<22} {'emails':>6} {'days':>5}")
    for r in recs:
        print(f"{r['needs']:<26} @{r['anchor']:<17} {r['from']:<22} {r['email_span']:>6} {r['day_span']:>5}")
    es = [r["email_span"] for r in recs]
    print(f"\nemail-span: max {max(es)}, mean {sum(es)/len(es):.1f}  "
          f"(spans this small => the fact is still in context; no retrieval forced)")


if __name__ == "__main__":
    args = sys.argv[1:]
    main(int(args[0]) if args else 42, int(args[1]) if len(args) > 1 else 100)
