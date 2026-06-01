# Run results — Haiku retrieval-span pilot (2026-05-31)

A plain-language record of the scaled Haiku run, what it showed, and what to watch out for.

## What we ran

- **Model:** claude-haiku-4-5, one continuous session (`claude -p --resume`), one email per turn.
- **Corpus:** 267 emails = the handwritten `corpus/` nodes + 200 long "junk" filler emails + 24 planted tests ("needles"), 6 in each of 4 difficulty tiers.
- **Reproduce it:**
  ```
  .venv/bin/python -m sb.scale --filler 200 --needles 6 --seed 42 --days 200
  NO_COLOR=1 ./run.sh --model claude-haiku-4-5 --corpus build/scaled --seed 42 --days 200 > build/haiku_tiered3.log 2>&1
  .venv/bin/python -m sb.analyze build/haiku_tiered3.log --corpus build/scaled --seed 42 --days 200
  ```

## The two dials we measured

- **Reasoning difficulty (tiers):** T1 = "one week after X" (easy) → T2 = "last business day before X" → T3 = "first Monday after X" → T4 = combine two facts from two emails (hard).
- **Span (memory pressure):** how many emails sit between the email that states a fact and the email that needs it. Small span = fact still in the assistant's memory; large span = it was trimmed away and the assistant must search the inbox to recover it.

## Headline result

**Needle accuracy: 86%** (21 of 24 needles; 3 were lost to a rate-limit window, see below). Haiku reasons well across all four tiers.

| difficulty | span 0–50 | span 50–100 | span 100+ |
|---|---|---|---|
| T1 simple offset | 50% (2) | 100% (1) | 100% (2) |
| T2 business-day | 100% (4) | 100% (1) | — |
| T3 weekday-after | 100% (2) | 100% (1) | 33% (3) |
| T4 multi-fact | 100% (1) | 100% (1) | 100% (3) |

| span | accuracy | how often it searched the inbox |
|---|---|---|
| 0–50 (fact still in memory) | 89% | 0% |
| 50–100 | 100% | 25% |
| 100+ (fact trimmed away) | 75% | 100% |

*(The number in parentheses is how many needles were in that cell — see caveat #1.)*

## What it means (plain)

1. **The model's date reasoning is good.** 86% across easy-to-hard math.
2. **It copes with memory pressure better than expected.** When a needed fact had been trimmed out of memory (span 100+), it searched its inbox **100% of the time** and still got **75%** right. It recognizes the gap and goes digging, rather than guessing.
3. **The earlier scary "21%" was our bug, not the model.** An earlier version of this run scored 21% because the answer key double-counted: the model would also create an event for the *announced* thing on the setup email, and our keyword match lumped it together with the correct answer. Fixed by matching the action word too (e.g. "relocation **review**", not just "relocation"). After the fix: 86%.

## Things to worry about / keep in mind

1. **The numbers are noisy — small samples.** Most grid cells have only 1–4 needles. A cell like "T3 100+ = 33%" is 1 of 3 — basically a coin flip, not a real measurement. **Do not trust individual cells yet.** Only the overall 86% and the span column (n=8–9) are even close to reliable. Fix = more needles + more seeds.
2. **Rate limits are a real operational risk.** A previous 685-email run mostly failed because we hit the Claude account usage cap (~2,000 calls in one evening). This run added retry-with-backoff, which cut errors from 370 → 35, but a sustained ~35-email rate window still slipped through and cost 3 needles. Big runs can still hit a hard usage cap that backoff can't outwait. Keep runs modest, or use a pay-as-you-go API key for large ones.
3. **Watch for "the model over-acts on announcements."** On FYI emails that announce a future event ("the migration is on Aug 3, no action needed"), Haiku tends to put that event on the calendar anyway. That's a real no-action-discipline weakness, and it's also what created the grading artifact above. Keep answer-key titles specific so it can't contaminate scoring.
4. **The benchmark doesn't strongly separate models yet at this difficulty.** Haiku scoring 86% means it isn't really struggling. To make it discriminating (paper-worthy), push harder: more multi-fact (T4) and multi-hop reasoning, larger spans, and enough needles/seeds for real statistics.
5. **The scaled corpus is throwaway.** It lives in `build/` (gitignored) and is regenerated from the command above. The seed (42) makes it reproducible. The handwritten corpus in `corpus/` is the permanent part.

## Status

Tooling is built and validated offline (45 unit tests pass; the "oracle" perfect-model solves the corpus 100%, which is our check that every answer is actually achievable). It is **not committed yet** — `sb/span.py`, `sb/scale.py`, `sb/analyze.py`, the grammar extension, the runner's backoff + `--corpus` flag, and the oracle title fix.
