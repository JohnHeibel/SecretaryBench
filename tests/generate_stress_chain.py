#!/usr/bin/env python3
"""Generate a stress-chain workbook so compaction can actually fire (FIX-10).

The default `Emails.xlsx` tops out at ~5-email chains (~15-25K tokens), well
under the 200K context window, so claude never compacts and Sprint 5's
"completes through compaction instead of crashing" claim is never demonstrated.

This builds a SEPARATE workbook (default `stress_chain.xlsx`) holding ONE
scenario with many emails, each carrying a large realistic body. Run with
conversation continuity on, the per-scenario session accumulates every turn, so
the context grows past the window and claude compacts mid-chain. The final email
carries a `CC-{date}` criterion so the chain still produces a score after
compaction.

Usage:
    python tests/generate_stress_chain.py                 # 60 emails, ~12k chars each
    python tests/generate_stress_chain.py --emails 80 --body-chars 16000 -o big.xlsx

Then, with the server up and claude authenticated:
    python engine.py stress_chain.xlsx
and watch for `COMPACTION fired` in the log and a non-zero score on the final
CC criterion. See SPRINT5_REMEDIATION FIX-10.
"""
from __future__ import annotations

import argparse

import pandas as pd

# Excel hard-caps a single cell at 32,767 characters (~8k tokens). Bodies are
# clamped to stay under it; reach the context window via MORE emails, not bigger
# bodies. At ~8k tokens/email you need ~25-30 emails to cross a 200k window.
_XLSX_CELL_MAX = 32000

COLUMNS = ["Scenario ID", "Scenario Type", "Email #", "Subject", "Body",
           "Sender", "Recipient(s)", "Success Criteria", "Puzzle Summary"]

# Realistic-ish business paragraphs so claude behaves as it would on real mail
# (not lorem filler). Cycled and lightly varied per email to reach the target
# size without being byte-identical (which would compress/cache trivially).
_PARAGRAPHS = [
    "Following yesterday's working session, the integration team confirmed that "
    "the data-migration dry run completed against the staging replica. Row counts "
    "reconciled within tolerance and the checksum job reported no drift on the "
    "primary fact tables. A handful of nullable columns still need backfill before "
    "we can promote the cutover plan to the change-advisory board.",
    "Finance flagged that the Q3 reforecast assumes the vendor consolidation lands "
    "on schedule. If the contract redlines slip past the end of the month, the "
    "savings move into Q4 and the headcount plan needs a corresponding revision. "
    "Procurement is holding a slot to walk through the revised terms.",
    "On reliability: the on-call rotation saw two paging events overnight, both "
    "traced to a noisy downstream dependency rather than our service. We added a "
    "circuit breaker and a tighter timeout; the error budget for the quarter is "
    "still healthy. A short postmortem is attached for the record.",
    "Marketing wants the launch narrative locked before the analyst briefings. The "
    "current draft leans on three customer proof points; legal has cleared two and "
    "is reviewing the third. We should decide whether to hold the briefing for the "
    "third quote or proceed with the two that are approved.",
    "Engineering capacity for the next sprint is constrained by the platform "
    "upgrade. We can either pull the upgrade forward and absorb a slower feature "
    "sprint, or defer it two weeks and accept the known deprecation warnings a "
    "little longer. Both paths are defensible; we need a call by Thursday.",
]


def _body(i: int, target_chars: int) -> str:
    """A realistic, per-email-varied body padded to ~target_chars characters."""
    head = (f"Project Atlas — running thread, message {i}.\n\n"
            f"Summary for today: this is update number {i} in the ongoing program "
            f"review. Please keep this thread in context; later messages refer back "
            f"to decisions logged here.\n\n")
    chunks = [head]
    n = 0
    para_idx = 0
    while sum(len(c) for c in chunks) < target_chars:
        p = _PARAGRAPHS[para_idx % len(_PARAGRAPHS)]
        chunks.append(f"[{i}.{n}] {p}\n\n")
        para_idx += 1
        n += 1
    return "".join(chunks)[:target_chars]


def generate(path: str, n_emails: int, body_chars: int) -> None:
    if body_chars > _XLSX_CELL_MAX:
        print(f"note: body-chars {body_chars} exceeds the Excel cell cap; clamping "
              f"to {_XLSX_CELL_MAX}. Use more --emails to reach the context window.")
        body_chars = _XLSX_CELL_MAX
    sid = "STRESS-CHAIN"
    stype = "STRESS01"
    rows = []
    for i in range(1, n_emails + 1):
        is_first = i == 1
        is_last = i == n_emails
        # The final email asks to book a wrap-up meeting -> a checkable criterion
        # that must still grade after compaction has fired earlier in the chain.
        criteria = "CC-{date}" if is_last else None
        subject = ("Project Atlas — schedule wrap-up review"
                   if is_last else f"Project Atlas update {i}/{n_emails}")
        body = _body(i, body_chars)
        if is_last:
            body += ("\n\nACTION: please put a 1-hour Project Atlas wrap-up review "
                     "on the calendar for today and confirm.")
        rows.append({
            "Scenario ID": sid if is_first else None,    # loader uses row-1 value
            "Scenario Type": stype if is_first else None,  # loader ffills downward
            "Email #": i,
            "Subject": subject,
            "Body": body,
            "Sender": "ops@bigco.example",
            "Recipient(s)": "assistant@bigco.example",
            "Success Criteria": criteria,
            "Puzzle Summary": ("Long-running project thread; book a wrap-up meeting "
                               "at the end.") if is_first else None,
        })

    df = pd.DataFrame(rows, columns=COLUMNS)
    df.to_excel(path, sheet_name="Stress", index=False)
    total_chars = sum(len(r["Body"]) for r in rows)
    print(f"Wrote {path}: 1 scenario, {n_emails} emails, ~{total_chars:,} body chars "
          f"(~{total_chars // 4:,} tokens of email body alone, before tool results).")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-o", "--out", default="stress_chain.xlsx")
    # 30 emails x ~8k tokens ≈ 240k accumulated context > a 200k window, so
    # compaction fires roughly two-thirds of the way through the chain.
    ap.add_argument("--emails", type=int, default=30)
    ap.add_argument("--body-chars", type=int, default=_XLSX_CELL_MAX)
    args = ap.parse_args()
    generate(args.out, args.emails, args.body_chars)
