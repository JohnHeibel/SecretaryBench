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

Before grading, the corpus is checked against the run (C-11). A corpus whose hash
differs from the manifest's is accepted only if the capture holds the frozen
corpus the run was served from (`<capture>/corpus`), the serve plan rebuilt from
the given corpus is the one the run used, and every captured email renders to
exactly what the model was served. Then only answer keys differ, and the grade
is a key-repair score against the frozen run. Anything else is refused: a prose
or plan repair needs a new run. `--corpus <capture>/corpus` grades against the
run's own keys.
"""
from __future__ import annotations

import argparse
import json
import sys
from dataclasses import asdict, dataclass, field
from datetime import date, timedelta
from pathlib import Path
from typing import Optional

from sb.live.runner import _corpus_hash, _grade_day, _secretary_status, _served_email
from sb.schema import Corpus, load_corpus
from sb.scheduler import InfeasibleSchedule, Levers, Plan, build_plan


class ReferenceMismatch(Exception):
    """The corpus would grade a capture against mail its model was not served, or on a
    schedule it was not run on (C-11). The message says what differs and what to do."""


@dataclass
class Reference:
    corpus_dir: str             # the corpus that will be graded with
    corpus_hash: str            # its hash
    run_hash: str               # the hash the run recorded
    identical: bool             # same corpus as the run
    keys_changed: list[str] = field(default_factory=list)   # captured emails whose answer differs


def load_capture(capture_dir: str) -> tuple[dict, list[dict]]:
    """Return (manifest, day records ordered by day number)."""
    d = Path(capture_dir)
    manifest = json.loads((d / "manifest.json").read_text())
    days = [json.loads(p.read_text()) for p in sorted((d / "days").glob("*.json"))]
    return manifest, days


def _plan(corpus: Corpus, manifest: dict) -> Plan:
    lv = manifest["levers"]
    return build_plan(
        corpus,
        start_date=date.fromisoformat(manifest["start"]),
        seed=manifest["seed"],
        n_days=manifest["n_days"],
        levers=Levers(daily_min=lv["daily_min"], daily_max=lv["daily_max"],
                      urgency_horizon=lv["urgency_horizon"]),
    )


def _listed(lines: list[str], limit: int = 10) -> str:
    more = f"\n  ... and {len(lines) - limit} more" if len(lines) > limit else ""
    return "\n".join(f"  {line}" for line in lines[:limit]) + more


def _plan_problems(manifest: dict, days: list[dict], plan: Plan) -> list[str]:
    """How the rebuilt plan differs from the one the capture was served on."""
    recorded = manifest["serve_plan"]
    out = [f"{eid}: served {recorded.get(eid)} in the run, {plan.serve_date.get(eid)} from this corpus"
           for rec in days for eid in rec["batch"]
           if recorded.get(eid) != str(plan.serve_date.get(eid))]

    # Batches in order, as the runner posts them; only --limit may cut the last day short.
    rebuilt = {str(plan.start_date + timedelta(days=i)): b for i, b in enumerate(plan.per_day) if b}
    for i, rec in enumerate(days):
        got, want = list(rec["batch"]), rebuilt.get(rec["serve_date"], [])
        cut_short = i == len(days) - 1 and manifest.get("limit")
        if want[:len(got)] != got or (len(want) != len(got) and not cut_short):
            out.append(f"day {rec['day']} ({rec['serve_date']}): the run served {got}, "
                       f"this corpus serves {want}")
    served_on = {rec["serve_date"] for rec in days}
    last = max(served_on, default="")
    out += [f"{d}: the run served nothing, this corpus serves {b}"
            for d, b in rebuilt.items() if d <= last and d not in served_on]
    return out


def check_reference(capture_dir: str, corpus_dir: Optional[str] = None) -> Reference:
    """Check that `corpus_dir` (default: the manifest's) can grade this capture: same
    corpus, or one that differs from the run's frozen corpus only in what the grader
    reads. Raises ReferenceMismatch otherwise (C-11)."""
    manifest, days = load_capture(capture_dir)
    given_dir = corpus_dir or manifest["corpus_dir"]
    given_hash, run_hash = _corpus_hash(given_dir), manifest.get("corpus_hash")
    if given_hash == run_hash:
        return Reference(given_dir, given_hash, run_hash, identical=True)

    head = (f"{capture_dir} was run on corpus sha {run_hash}; "
            f"{given_dir} is sha {given_hash}.")
    snap = Path(capture_dir) / "corpus"
    remedy = f"pass --corpus {snap} to grade against the run's own corpus"
    if not any((snap / "nodes").glob("*.json")):
        raise ReferenceMismatch(
            f"{head}\ncannot verify: no frozen corpus in the capture ({snap}/nodes).\n"
            f"pass --corpus with the corpus the run used (sha {run_hash}) to grade against "
            f"the run's own corpus")
    snap_hash = _corpus_hash(str(snap))
    if snap_hash != run_hash:
        raise ReferenceMismatch(
            f"{head}\nfrozen corpus does not match the run: {snap} is sha {snap_hash}.\n"
            f"restore {snap} from version control, then {remedy}")

    frozen, given = load_corpus(snap), load_corpus(given_dir)
    captured = [eid for rec in days for eid in rec["batch"]]
    removed = [e for e in frozen.emails if e not in given.emails]
    added = [e for e in given.emails if e not in frozen.emails]
    if removed or added:
        raise ReferenceMismatch(
            f"{head}\nthe corpus has different emails from the run's:\n"
            + _listed([f"{e}: removed" + (" (captured)" if e in captured else "") for e in removed]
                      + [f"{e}: added" for e in added])
            + f"\n{remedy}")

    if "serve_plan" not in manifest:
        raise ReferenceMismatch(f"{head}\ncannot verify: the manifest records no serve plan.\n{remedy}")
    try:
        plan = _plan(given, manifest)
    except InfeasibleSchedule as exc:
        raise ReferenceMismatch(f"{head}\nits serve plan does not build: {exc}\n{remedy}") from None
    problems = _plan_problems(manifest, days, plan)
    if problems:
        raise ReferenceMismatch(
            f"{head}\nit changes the serve plan the run was served on "
            f"(seed {manifest['seed']}, start {manifest['start']}, levers {manifest['levers']}):\n"
            + _listed(problems) + f"\n{remedy}")

    frozen_plan = _plan(frozen, manifest)
    changed = []
    for eid in captured:
        now, was = _served_email(given, plan, eid), _served_email(frozen, frozen_plan, eid)
        if now != was:
            changed.append(f"{eid}: {', '.join(k for k in was if now[k] != was[k])}")
    if changed:
        raise ReferenceMismatch(
            f"{head}\nit changes what the model was served on {len(changed)} email(s):\n"
            + _listed(changed) + f"\n{remedy}")

    keys = [eid for eid in captured if given.emails[eid].answer != frozen.emails[eid].answer]
    return Reference(given_dir, given_hash, run_hash, identical=False, keys_changed=keys)


def regrade(capture_dir: str, corpus_dir: Optional[str] = None) -> dict:
    """Re-run the grader over a captured run. Returns {email_id: EmailResult|None}.
    Raises ReferenceMismatch if the corpus cannot grade this run (C-11)."""
    check_reference(capture_dir, corpus_dir)
    manifest, days = load_capture(capture_dir)
    corpus = load_corpus(corpus_dir or manifest["corpus_dir"])
    plan = _plan(corpus, manifest)

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
                    help="override the corpus recorded in the manifest (checked against "
                         "the run; <capture>/corpus is the run's own)")
    ap.add_argument("--json", action="store_true", help="emit per-email verdicts as JSON")
    a = ap.parse_args()

    manifest, days = load_capture(a.capture)
    try:
        ref = check_reference(a.capture, a.corpus)
    except ReferenceMismatch as exc:
        print(f"REFUSED   {exc}", file=sys.stderr)
        raise SystemExit(2)
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
    n = len(ref.keys_changed)
    same = ("(same as the run)" if ref.identical else
            f"(run: {ref.run_hash}; served emails and plan identical; "
            f"answer keys differ on {n} email{'' if n == 1 else 's'})")
    print(f"corpus    {ref.corpus_dir}  sha {ref.corpus_hash}  {same}")
    if n:
        print(f"          {', '.join(ref.keys_changed[:10])}"
              + (f" and {n - 10} more" if n > 10 else ""))
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
