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

It prints two scores. RESCORE counts every email. VALID drops the emails on
days that did not measure the model: days the runner marked errored, and days
whose turn started without the secretary MCP server connected (register O-7).
Captures written before O-7 record the latter only in the CLI's init event, so
it is read from raw/day_NNN.jsonl; day records are never edited.
"""
from __future__ import annotations

import argparse
import json
from dataclasses import asdict
from datetime import date
from pathlib import Path
from typing import Optional

from sb.live.runner import _grade_day, _secretary_status
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


def mcp_status(capture_dir: str, rec: dict) -> Optional[str]:
    """The secretary server's status when the day's graded turn started, or None if
    unobservable (codex driver, or no raw stream)."""
    if "mcp_status" in rec:                      # recorded by the runner since O-7
        return rec["mcp_status"]
    raw = Path(capture_dir) / "raw" / f"day_{rec['day']:03d}.jsonl"
    return _secretary_status(raw.read_text()) if raw.exists() else None


def invalid_days(capture_dir: str, days: list[dict]) -> dict[int, str]:
    """{day: reason} for every day that does not measure the model."""
    out = {}
    for rec in days:
        if not rec.get("ok"):
            out[rec["day"]] = "errored"
            continue
        status = mcp_status(capture_dir, rec)
        if status is not None and status != "connected":
            out[rec["day"]] = f"secretary MCP {status}"
    return out


def _emails_on(days: list[dict], bad: dict[int, str]) -> dict[str, str]:
    return {eid: bad[rec["day"]] for rec in days if rec["day"] in bad for eid in rec["batch"]}


def validity(capture_dir: str) -> dict[str, str]:
    """{email_id: reason} for every email on an invalid day."""
    _, days = load_capture(capture_dir)
    return _emails_on(days, invalid_days(capture_dir, days))


def score_valid(results: dict, invalid: dict) -> tuple[int, int]:
    """(passed, total) over the emails on valid days only."""
    valid = [r for eid, r in results.items() if eid not in invalid]
    return sum(1 for r in valid if r is not None and r.passed), len(valid)


def _day_ranges(days: list[int]) -> str:
    runs: list[list[int]] = []
    for d in sorted(days):
        if runs and d == runs[-1][-1] + 1:
            runs[-1].append(d)
        else:
            runs.append([d])
    return ", ".join(f"{r[0]}-{r[-1]}" if len(r) > 1 else f"{r[0]}" for r in runs)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    ap.add_argument("capture", help="a directory written by `sb.live.runner --out`")
    ap.add_argument("--corpus", default=None,
                    help="override the corpus recorded in the manifest")
    ap.add_argument("--json", action="store_true", help="emit per-email verdicts as JSON")
    a = ap.parse_args()

    manifest, days = load_capture(a.capture)
    results = regrade(a.capture, a.corpus)
    passed, total, errored = score(results)

    if a.json:
        print(json.dumps({eid: (asdict(r) if r else None)
                          for eid, r in results.items()}, indent=2, default=str))
        return

    bad_days = invalid_days(a.capture, days)
    invalid = _emails_on(days, bad_days)
    v_passed, v_total = score_valid(results, invalid)
    unobserved = sum(1 for rec in days if rec.get("ok") and mcp_status(a.capture, rec) is None)

    served = manifest.get("served_model") or manifest.get("requested_model")
    cert = "certified" if manifest.get("model_certified") else "ASSERTED, not observed"
    print(f"capture   {a.capture}")
    print(f"model     {served}  ({cert})")
    print(f"corpus    {a.corpus or manifest['corpus_dir']}  sha {manifest.get('corpus_hash')}")
    print(f"levers    {manifest['levers']}  seed {manifest['seed']}")
    print(f"RESCORE   {passed}/{total} ({passed / total:.0%})  all days"
          + (f", {errored} errored count as failed" if errored else ""))
    by_reason: dict[str, list[int]] = {}
    for d, why in bad_days.items():
        by_reason.setdefault(why, []).append(d)
    excluded = "; ".join(f"days {_day_ranges(ds)} {r}" for r, ds in by_reason.items())
    print(f"VALID     {v_passed}/{v_total} ({v_passed / v_total if v_total else 0:.0%})  "
          + (f"{len(invalid)} emails excluded: {excluded}" if bad_days else "every day valid"))
    if unobserved:
        print(f"          {unobserved} day(s) have no secretary MCP status on record, counted valid")
    was = manifest.get("score_passed")
    if was is not None:
        delta = passed - was
        sign = "+" if delta > 0 else ""
        print(f"as-run    {was}/{manifest.get('score_total')}"
              + (f"   delta {sign}{delta}" if delta else "   (identical)"))


if __name__ == "__main__":
    main()
