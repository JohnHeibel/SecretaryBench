"""
sb.analyze — join a live run log with offline span to answer the paper's question:
does a needle's correctness fall as the fact it needs sits further back (higher
span, more likely compacted out of context)?

Reads the runner's printed log (NO_COLOR), pairs each email with PASS/FAIL and
whether the model searched the inbox, and recomputes span on the same scaled
corpus/seed. Prints a 2D reasoning-tier × span-bin accuracy grid plus search-rate
summary.

    python -m sb.analyze build/haiku_run.log --corpus build/scaled --seed 42 --days 300
"""
from __future__ import annotations

import argparse
import json
import os
import re
from datetime import date

from sb.scheduler import build_plan
from sb.schema import load_corpus
from sb.span import spans

_LINE = re.compile(r"\b(PASS|FAIL)\b\s+\[(\d+)\]\s+(\S+)")

SPAN_BINS = [(0, 50), (50, 100), (100, 10**9)]
BIN_LABELS = ["0-50", "50-100", "100+"]
TIER_ORDER = ["T1", "T2", "T3", "T4"]


def parse_log(path: str) -> dict[str, dict]:
    """email_id -> {passed, searched}. 'searched' read from the tools line."""
    out: dict[str, dict] = {}
    cur = None
    for raw in open(path, encoding="utf-8", errors="replace"):
        m = _LINE.search(raw)
        if m:
            cur = m.group(3)
            out[cur] = {"passed": m.group(1) == "PASS", "searched": False}
        elif cur and "tools" in raw:
            out[cur]["searched"] = ("inbox" in raw or "get_email" in raw)
            cur = None
    return out


def _bin(span: int) -> int:
    for i, (lo, hi) in enumerate(SPAN_BINS):
        if lo <= span < hi:
            return i
    return len(SPAN_BINS) - 1


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("log")
    ap.add_argument("--corpus", default="build/scaled")
    ap.add_argument("--seed", type=int, default=42)
    ap.add_argument("--days", type=int, default=300)
    a = ap.parse_args()

    results = parse_log(a.log)
    corpus = load_corpus(a.corpus)
    plan = build_plan(corpus, start_date=date(2026, 6, 1), seed=a.seed, n_days=a.days)
    span_of = {r["needs"]: r["email_span"] for r in spans(corpus, plan)}

    # --- load needle manifest if present ---
    manifest_path = os.path.join(a.corpus, "needles.json")
    rows: list[dict] = []
    if os.path.exists(manifest_path):
        manifest = json.load(open(manifest_path, encoding="utf-8"))
        for entry in manifest:
            pid = entry["payoff_id"]
            if pid not in span_of or pid not in results:
                continue
            rows.append({
                "payoff_id": pid,
                "tier": entry.get("reasoning_tier", "untagged"),
                "tier_name": entry.get("tier_name", ""),
                "span": span_of[pid],
                "passed": results[pid]["passed"],
                "searched": results[pid]["searched"],
            })
    else:
        # no manifest — treat every span record whose payoff is in results as "untagged"
        for eid, sp in span_of.items():
            if eid not in results:
                continue
            rows.append({
                "payoff_id": eid,
                "tier": "untagged",
                "tier_name": "untagged",
                "span": sp,
                "passed": results[eid]["passed"],
                "searched": results[eid]["searched"],
            })

    if not rows:
        print("no needle results found in log (run may be incomplete).")
        return

    # sort: tier order first, then span
    tier_key = lambda t: (TIER_ORDER.index(t) if t in TIER_ORDER else len(TIER_ORDER), t)
    rows.sort(key=lambda r: (tier_key(r["tier"]), r["span"]))

    # ── 1. Per-needle table ────────────────────────────────────────────────────
    print(f"\n{'payoff_id':<32} {'tier':>5} {'span':>5} {'result':>7} {'searched':>9}")
    print("-" * 62)
    for r in rows:
        print(f"{r['payoff_id']:<32} {r['tier']:>5} {r['span']:>5} {'PASS' if r['passed'] else 'FAIL':>7} {'yes' if r['searched'] else 'no':>9}")

    # ── 2. 2D grid: tier × span-bin ───────────────────────────────────────────
    tiers = sorted({r["tier"] for r in rows}, key=lambda t: tier_key(t))
    # cell data: grid[tier][bin_idx] = [passed_count, total_count]
    grid: dict[str, list[list[int]]] = {t: [[0, 0] for _ in BIN_LABELS] for t in tiers}
    for r in rows:
        bi = _bin(r["span"])
        grid[r["tier"]][bi][1] += 1
        if r["passed"]:
            grid[r["tier"]][bi][0] += 1

    col_w = 12
    tier_w = 10
    header = f"\n{'tier':<{tier_w}}" + "".join(f"{lbl:^{col_w}}" for lbl in BIN_LABELS)
    print(header)
    print("-" * (tier_w + col_w * len(BIN_LABELS)))
    for t in tiers:
        row_str = f"{t:<{tier_w}}"
        for bi in range(len(BIN_LABELS)):
            passed, n = grid[t][bi]
            if n == 0:
                cell = "·"
            else:
                cell = f"{passed/n:.0%} ({n})"
            row_str += f"{cell:^{col_w}}"
        print(row_str)

    # ── 3. Search-rate summary by span bin ────────────────────────────────────
    print(f"\n{'span bin':<12} {'n':>4} {'accuracy':>9} {'searched%':>10}")
    print("-" * 38)
    for bi, lbl in enumerate(BIN_LABELS):
        grp = [r for r in rows if _bin(r["span"]) == bi]
        if not grp:
            continue
        acc = sum(1 for r in grp if r["passed"]) / len(grp)
        srch = sum(1 for r in grp if r["searched"]) / len(grp)
        print(f"{lbl:<12} {len(grp):>4} {acc:>8.0%} {srch:>9.0%}")

    # ── 4. Overall ────────────────────────────────────────────────────────────
    overall = sum(1 for r in rows if r["passed"]) / len(rows)
    print(f"\noverall needle accuracy: {overall:.0%} over {len(rows)} needles")


if __name__ == "__main__":
    main()
