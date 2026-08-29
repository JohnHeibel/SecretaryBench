"""
sb.regrade — re-score a captured run offline, without re-running the model.

This is the point of `sb.live.runner --out`. A capture holds, per day, the FULL
store state the grader saw: every object the model created, with title,
description and attribution, including the ones matching no answer-key keyword.
So any change to `sb.grader` can be evaluated against a run that has already
been paid for, for free, as many times as you like.

Before this existed, evaluating a grader change meant paying for a whole new
run (register O-1). Note the boundary: this can re-score any rule, including
ones that newly admit objects the printed log never rendered, because the
capture keeps the objects rather than the log's rendering of them.

    python -m sb.regrade build/capture_x
    python -m sb.regrade build/capture_x --corpus corpus
"""
from __future__ import annotations

import argparse
import json
from dataclasses import asdict
from datetime import date
from pathlib import Path
from typing import Optional

from sb.live.runner import _grade_day
from sb.schema import load_corpus
from sb.scheduler import Levers, build_plan


def load_capture(capture_dir: str) -> tuple[dict, list[dict]]:
    """Return (manifest, day records ordered by day number)."""
    d = Path(capture_dir)
    manifest = json.loads((d / "manifest.json").read_text())
    days = [json.loads(p.read_text()) for p in sorted((d / "days").glob("*.json"))]
    return manifest, days


def regrade(capture_dir: str, corpus_dir: Optional[str] = None) -> dict:
    """Re-run the grader over a captured run. Returns {email_id: EmailResult|None}."""
    manifest, days = load_capture(capture_dir)
    corpus = load_corpus(corpus_dir or manifest["corpus_dir"])
    lv = manifest["levers"]
    plan = build_plan(
        corpus,
        start_date=date.fromisoformat(manifest["start"]),
        seed=manifest["seed"],
        n_days=manifest["n_days"],
        levers=Levers(daily_min=lv["daily_min"], daily_max=lv["daily_max"],
                      urgency_horizon=lv["urgency_horizon"]),
    )

    results: dict[str, object] = {}
    for rec in days:
        if not rec.get("ok"):
            for eid in rec["batch"]:
                results[eid] = None
            continue
        results.update(_grade_day(corpus, plan, list(rec["batch"]), rec["state"], set(rec["day_new"])))
    return results


def score(results: dict) -> tuple[int, int, int]:
    """(passed, total, errored)"""
    graded = [r for r in results.values() if r is not None]
    return sum(1 for r in graded if r.passed), len(results), len(results) - len(graded)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    ap.add_argument("capture", help="a directory written by `sb.live.runner --out`")
    ap.add_argument("--corpus", default=None,
                    help="override the corpus recorded in the manifest")
    ap.add_argument("--json", action="store_true", help="emit per-email verdicts as JSON")
    a = ap.parse_args()

    manifest, _ = load_capture(a.capture)
    results = regrade(a.capture, a.corpus)
    passed, total, errored = score(results)

    if a.json:
        print(json.dumps({eid: (asdict(r) if r else None)
                          for eid, r in results.items()}, indent=2, default=str))
        return

    served = manifest.get("served_model") or manifest.get("requested_model")
    cert = "certified" if manifest.get("model_certified") else "ASSERTED, not observed"
    print(f"capture   {a.capture}")
    print(f"model     {served}  ({cert})")
    print(f"corpus    {a.corpus or manifest['corpus_dir']}  sha {manifest.get('corpus_hash')}")
    print(f"levers    {manifest['levers']}  seed {manifest['seed']}")
    print(f"RESCORE   {passed}/{total} ({passed / total:.0%})"
          + (f"  {errored} errored" if errored else ""))
    was = manifest.get("score_passed")
    if was is not None:
        delta = passed - was
        sign = "+" if delta > 0 else ""
        print(f"as-run    {was}/{manifest.get('score_total')}"
              + (f"   delta {sign}{delta}" if delta else "   (identical)"))


if __name__ == "__main__":
    main()
