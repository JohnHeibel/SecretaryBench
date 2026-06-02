"""
sb.scale — build a SCALED corpus to drive retrieval span, and measure it offline.

The base corpus keeps every needed fact in context (see `python -m sb.span`), so a
model rarely has to retrieve. This tool copies the real corpus into a build dir and
injects a haystack of no-action FYI "filler" emails — the distractors that push the
facts out of context and that search_inbox must sift through.

The needles themselves are NOT made here. A needle is an authored pair already in
corpus/: an early email that emits a date anchor ({!name = ...}) and a later email
whose answer key needs that anchor (@name). This tool just slips filler BETWEEN them
so the span (gap from setup to payoff) grows until the fact scrolls out of context.
More --filler => larger span => more retrieval forced.

Everything here is OFFLINE (no API calls): copy -> filler -> plan -> span -> oracle.
The real corpus/ and the test suite are untouched (we build into build/scaled/).

    python -m sb.scale --filler 200 --seed 42      # bury the authored needles in junk
"""
from __future__ import annotations

import argparse
import json
import random
import shutil
from datetime import date
from pathlib import Path

from sb.engine import Store, run
from sb.oracle import oracle_model
from sb.schema import load_corpus
from sb.scheduler import build_plan
from sb.span import spans

REPO = Path(__file__).resolve().parents[1]

_SENDERS = ["IT Helpdesk", "HR Bulletin", "Facilities", "Payroll", "Travel Desk",
            "Security Team", "Office Manager", "Vendor Updates", "All-Hands", "Benefits"]
_SUBJECTS = ["Weekly newsletter", "System maintenance notice", "Policy FYI",
             "Reminder: badge renewal", "Survey results are in", "Office closure note",
             "Software update available", "Quarterly digest", "Parking advisory",
             "Wellness tip of the week", "Receipt confirmation", "Mailbox almost full"]
# Paragraph pool for realistically LONG no-action filler. Real business email is
# paragraphs, not one line; long filler overflows the context window at a few
# hundred emails instead of thousands, so brute-scale overflow stays affordable.
# Every paragraph is explicitly informational so the email needs no action.
_PARAS = [
    "Thanks for your continued patience as we work through the backlog of operational updates this quarter. The team has been heads-down on a number of behind-the-scenes improvements, and we wanted to make sure everyone has visibility even though nothing here requires a response.",
    "As a reminder, all of the figures referenced below are preliminary and provided purely for situational awareness. They will be reconciled during the normal close process, and you do not need to take any action or reply to this message.",
    "We reviewed the usual dashboards and everything is tracking within expected ranges. A few minor anomalies were noted and have already been routed to the appropriate owners, so there is nothing outstanding that needs your attention at this time.",
    "Per our standard cadence, this notice is being distributed to the broader stakeholder list. If any of the items become relevant to your area, the responsible team will reach out directly; otherwise you can treat this as informational only.",
    "The vendor confirmed that the previously communicated timelines remain unchanged. We will continue to monitor and will only escalate if something materially shifts. No calendar changes or approvals are needed from you in the meantime.",
    "For completeness we are including the recap from the last working session. It captures decisions that were already made and ratified, so it is shared here strictly for the record and does not require further discussion or scheduling.",
    "Several people asked for a consolidated summary, so we have folded the highlights into this update. As always, this is a one-way broadcast and there is no expectation of a reply, acknowledgment, or follow-up meeting.",
    "Finally, a gentle note that our communication preferences page lets you adjust the frequency of these digests. Nothing about your current settings needs to change, and this message itself is purely a courtesy notification.",
    "Attendance figures from the optional sessions came in slightly above forecast, which the organizing committee is treating as a positive signal. No decisions hinge on these numbers, and they are shared only so the broader group has the same context as the planning team.",
    "We also want to acknowledge the volume of cross-functional traffic this period generated. Threads have been archived in the usual shared location for anyone who wants to browse, but none of them contain open items assigned to you or require sign-off of any kind.",
    "On the systems side, the routine patch window completed without incident and all services returned to nominal well inside the maintenance envelope. There is no expected user-visible impact, and we are noting it here purely for the change-log record.",
    "A handful of recurring questions have been added to the internal knowledge base so future newsletters can stay short. This is housekeeping on our end and does not imply any change to your workflows, approvals, or calendar.",
    "The finance team reminds everyone that the figures circulating informally are not the figures of record until the cycle closes. We mention it only to prevent confusion; there is nothing to reconcile, approve, or schedule on your part right now.",
    "As we wrap up, thank you again for reading to the end of these comprehensive updates. Consolidating everything into one message is meant to reduce inbox noise, and as always nothing in this digest is actionable or time-sensitive for you.",
    "For transparency we are also restating the standing guidance that informational broadcasts like this one never carry deadlines. If a genuine action is ever needed, it will arrive as a separate, clearly-marked request addressed directly to you.",
    "The working group rotated a few responsibilities internally to balance load this quarter. The hand-offs are fully handled within the team, the directory has been updated, and there is nothing you need to do or acknowledge in response.",
]


def _filler_node(n: int, rng: random.Random) -> dict:
    emails = []
    for i in range(1, n + 1):
        k = rng.randint(16, 22)                    # long, realistic emails => context overflows
        body = "\n\n".join(rng.choices(_PARAS, k=k))   # at ~150 emails, not ~600
        emails.append({
            "id": f"gen.filler.{i:04d}",
            "from": rng.choice(_SENDERS),
            "to": "CEO",
            "subject": f"{rng.choice(_SUBJECTS)} #{i}",
            "body": body,
            "depends_on": [],
            "answer": {"ops": []},
        })
    return {"id": "gen-filler", "cast": {"CEO": "you"}, "emails": emails}


def build_scaled(dst: Path, n_filler: int, seed: int) -> None:
    """Copy the authored corpus into dst/ and add one node of n_filler junk emails.
    The needles are whatever the authored corpus already contains; we only add the
    haystack that spaces their setup and payoff apart."""
    nodes = dst / "nodes"
    if dst.exists():
        shutil.rmtree(dst)
    nodes.mkdir(parents=True)
    for f in (REPO / "corpus" / "nodes").glob("*.json"):
        shutil.copy(f, nodes / f.name)
    rng = random.Random(seed)
    (nodes / "gen_filler.json").write_text(json.dumps(_filler_node(n_filler, rng), indent=2))


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--filler", type=int, default=120, help="no-action junk emails to bury the authored needles under")
    ap.add_argument("--seed", type=int, default=42)
    ap.add_argument("--days", type=int, default=200)
    ap.add_argument("--dst", default="build/scaled")
    a = ap.parse_args()

    dst = REPO / a.dst
    build_scaled(dst, a.filler, a.seed)
    corpus = load_corpus(str(dst))
    plan = build_plan(corpus, start_date=date(2026, 6, 1), seed=a.seed, n_days=a.days)
    served = [e for b in plan.per_day for e in b]

    recs = sorted(spans(corpus, plan), key=lambda r: r["email_span"], reverse=True)
    print(f"\nscaled corpus: {len(corpus.emails)} emails, {len(served)} served over {plan.n_days} days")
    print(f"filler: {a.filler} junk emails burying {len(recs)} authored needle(s)")
    es = [r["email_span"] for r in recs]
    if es:
        print(f"needle span: max {max(es)}, mean {sum(es)/len(es):.1f}  (n={len(es)})")
    else:
        print("needle span: no answer-key anchor references in this corpus")

    store = Store(corpus)
    res = run(corpus, plan, oracle_model, store=store)
    print(f"oracle: {res.passed}/{res.total} = {res.score():.0%} (must be 100% — corpus is valid at scale)\n")


if __name__ == "__main__":
    main()
