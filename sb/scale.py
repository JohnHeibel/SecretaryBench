"""
sb.scale — build a SCALED corpus to drive retrieval span, and measure it offline.

The base corpus keeps every needed fact in context (see `python -m sb.span`), so a
model never has to retrieve. This tool copies the real corpus into a build dir and
injects a haystack of no-action FYI "filler" emails — the distractors that push
facts out of context and that search_inbox must sift through.

The DISCRIMINATING needles (an early email emitting a date anchor + a later payoff
whose answer needs it with a far deadline) are hand-authored in corpus/ via the web
app, NOT generated — templated needles don't separate models. So the default run is
JUNK-ONLY (--needles 0): it buries the authored needles in haystack. --needles N>0
additionally injects N generated needles per reasoning tier, for span experiments.

Everything here is OFFLINE (no API calls): generate -> plan -> span -> oracle. The
real corpus/ and the test suite are untouched (we build into build/scaled/).

    python -m sb.scale --filler 200 --seed 42      # junk-only: bury authored needles
    python -m sb.scale --filler 120 --needles 8    # + generated needles (span study)
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
            "answer": {"expect": []},
        })
    return {"id": "gen-filler", "cast": {"CEO": "you"}, "emails": emails}


# Distinct, searchable topics; each keyword is a lowercased word that appears in
# its topic phrase, so a natural event title ("Datacenter migration review") always
# contains it (no "pentest" vs "pen test" false negatives).
_TOPICS = [
    ("datacenter migration", "migration"), ("vendor security audit", "audit"),
    ("quarterly budget review", "budget"), ("office relocation", "relocation"),
    ("brand relaunch", "relaunch"), ("payroll system cutover", "cutover"),
    ("warehouse inventory count", "inventory"), ("board offsite", "offsite"),
    ("compliance training rollout", "compliance"), ("CRM upgrade", "upgrade"),
    ("supplier renegotiation", "supplier"), ("data retention purge", "retention"),
    ("network failover drill", "failover"), ("benefits open enrollment", "enrollment"),
    ("annual pricing reset", "pricing"), ("customer summit", "summit"),
    ("warehouse lease signing", "lease"), ("API deprecation", "deprecation"),
    ("security penetration test", "penetration"), ("fiscal year close", "fiscal"),
    ("product launch", "launch"), ("hiring freeze review", "freeze"),
    ("server decommission", "decommission"), ("tax filing", "tax"),
    ("insurance renewal", "insurance"), ("trademark registration", "trademark"),
    ("facilities inspection", "inspection"), ("disaster recovery test", "recovery"),
    ("marketing campaign", "campaign"), ("sales conference", "conference"),
    ("investor briefing", "investor"), ("patent review", "patent"),
    ("logistics overhaul", "logistics"), ("onboarding revamp", "onboarding"),
    ("procurement freeze", "procurement"), ("analytics rollout", "analytics"),
    ("cloud migration", "cloud"), ("partnership signing", "partnership"),
    ("audit remediation", "remediation"), ("capacity planning", "capacity"),
]

_TIER_NAME = {"T1": "simple-offset", "T2": "business-day",
              "T3": "anchor-weekday", "T4": "multi-fact"}
_NUM = {2: "two", 3: "three", 4: "four"}
_TIERS = ["T1", "T2", "T3", "T4"]


def _needle(tier: str, idx: int, topic: str, kw: str, rng: random.Random) -> tuple[dict, dict]:
    """One needle node + its manifest entry. Every tier states the event DATE only
    in the SETUP email; the payoff requires a computation the payoff email does NOT
    spell out, so the model must retrieve the fact (high span) and reason:
      T1 simple offset    -> @date + 1 week
      T2 business-day     -> last business day before @date (crosses weekends)
      T3 anchor-weekday   -> first Monday after @date (uses `next:MON from @date`)
      T4 multi-fact       -> @date + an interval stated in a SEPARATE policy email
    """
    weeks = rng.choice([4, 5, 6, 7, 8, 9, 10])     # setup date offset -> varies span
    anc = f"d{tier.lower()}{idx:02d}"
    base = f"gen.needle.{tier.lower()}.{idx:02d}"
    setup = {"id": f"{base}.setup", "from": "OPS", "to": "CEO",
             "subject": f"{topic.capitalize()} date confirmed",
             "body": f"For the record: the {topic} is scheduled for {{!{anc} = serve+{weeks}w}}. No action needed yet.",
             "depends_on": [], "answer": {"expect": []}}
    deps = [{"email": f"{base}.setup", "type": "date"}]
    emails = [setup]

    # `noun` is the distinctive action word in the payoff title; including it in
    # title_match (alongside the topic kw) means the count:1 check matches only the
    # payoff's event, not any stray event the model created from the setup announcement.
    if tier == "T1":
        instr, start, noun = f"Please put the {topic} review on the calendar for one week after the {topic}.", f"@{anc}+1w", "review"
    elif tier == "T2":
        instr, start, noun = f"Please put the {topic} go/no-go review on the calendar for the last business day before the {topic}.", f"@{anc}-1bd", "review"
    elif tier == "T3":
        instr, start, noun = f"Please put the {topic} kickoff on the calendar for the first Monday after the {topic}.", f"next:MON from @{anc}", "kickoff"
    else:  # T4 — the interval lives in a SEPARATE policy email the model must also find
        r = rng.choice([2, 3, 4])
        emails.append({"id": f"{base}.policy", "from": "OPS", "to": "CEO",
                       "subject": "Operations playbook reminder",
                       "body": f"Reminder from the operations playbook: the retrospective for the {topic} is always held {_NUM[r]} weeks after the {topic} itself. Filing for reference; no action needed.",
                       "depends_on": [], "answer": {"expect": []}})
        deps.append({"email": f"{base}.policy", "type": "static"})
        instr, start, noun = f"Please schedule the {topic} retrospective per the interval in the operations playbook.", f"@{anc}+{r}w", "retrospective"

    emails.append({"id": f"{base}.payoff", "from": "OPS", "to": "CEO",
                   "subject": f"{topic.capitalize()} follow-up", "body": instr,
                   "depends_on": deps,
                   "answer": {"expect": [{"action": "create_event", "title_match": [kw, noun],
                                          "start": {"eq": start}, "count": 1}]}})
    node = {"id": f"gen-needle-{tier.lower()}-{idx:02d}",
            "cast": {"OPS": "Operations", "CEO": "you"}, "emails": emails}
    entry = {"payoff_id": f"{base}.payoff", "reasoning_tier": tier,
             "tier_name": _TIER_NAME[tier], "anchor": anc}
    return node, entry


def _needle_nodes(n_per_tier: int, rng: random.Random) -> tuple[list[dict], list[dict]]:
    nodes, manifest, gid = [], [], 0
    for tier in _TIERS:
        for idx in range(n_per_tier):
            topic, kw = _TOPICS[gid % len(_TOPICS)]
            gid += 1
            node, entry = _needle(tier, idx, topic, kw, rng)
            nodes.append(node)
            manifest.append(entry)
    return nodes, manifest


def build_scaled(dst: Path, n_filler: int, seed: int, n_per_tier: int = 8) -> None:
    nodes = dst / "nodes"
    if dst.exists():
        shutil.rmtree(dst)
    nodes.mkdir(parents=True)
    for f in (REPO / "corpus" / "nodes").glob("*.json"):
        shutil.copy(f, nodes / f.name)
    rng = random.Random(seed)
    (nodes / "gen_filler.json").write_text(json.dumps(_filler_node(n_filler, rng), indent=2))
    needle_nodes, manifest = _needle_nodes(n_per_tier, rng)
    for nd in needle_nodes:
        (nodes / f"{nd['id'].replace('-', '_')}.json").write_text(json.dumps(nd, indent=2))
    (dst / "needles.json").write_text(json.dumps(manifest, indent=2))


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--filler", type=int, default=120)
    # Default 0: the discriminating needles are hand-authored in the corpus (via the
    # web app), not generated. Generated needles are templated and don't separate
    # models, so the production use of this tool is junk-only — bury the real,
    # authored needles in haystack. Pass --needles N>0 only for span experiments.
    ap.add_argument("--needles", type=int, default=0, help="GENERATED needles per reasoning tier (x4); 0 = junk-only around the authored corpus")
    ap.add_argument("--seed", type=int, default=42)
    ap.add_argument("--days", type=int, default=200)
    ap.add_argument("--dst", default="build/scaled")
    a = ap.parse_args()

    dst = REPO / a.dst
    build_scaled(dst, a.filler, a.seed, a.needles)
    corpus = load_corpus(str(dst))
    plan = build_plan(corpus, start_date=date(2026, 6, 1), seed=a.seed, n_days=a.days)
    served = [e for b in plan.per_day for e in b]

    recs = sorted(spans(corpus, plan), key=lambda r: r["email_span"], reverse=True)
    print(f"\nscaled corpus: {len(corpus.emails)} emails, {len(served)} served over {plan.n_days} days")
    kind = f"{a.needles} per tier x4 ({', '.join(_TIERS)})" if a.needles else "0 generated — burying the authored corpus needles in junk"
    print(f"needles: {kind}")
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
