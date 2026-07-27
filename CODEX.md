# CODEX.md — task brief for the corpus-authoring pass

You are working in **SecretaryBench**, an academic temporal-reasoning benchmark for LLM
"secretary" agents. The harness code (`sb/`) is done and merged. **Your job is the email
corpus** in `corpus/nodes/*.json`: get the current drafts to lint clean, pass the oracle at
100%, and carry correct difficulty tiers.

Do NOT touch `sb/` engine code unless a genuine bug blocks you (flag it separately if so).

---

## Read these first (in order)

1. `RECAP.md` — the system overview.
2. `ANSWER_KEY_GRAMMAR.md` — the exact date-token + answer-key (`ops`) grammar. This is law.
3. `TIER_LIST.md` — the T1/T2/T3 authoring playbook (what makes an email easy vs hard).
4. `RUNNING.md` — how the corpus is scaled, served, and graded.
5. `sb/tests/fixtures/nodes/{alpha,beta,gamma}.json` — the **canonical format reference**.
   When in doubt about node shape, copy these.

---

## Current state (verify it yourself)

`corpus/nodes/` holds ~23 untracked draft nodes of mixed quality. Known problems to expect:

- **Empty stubs**: `node-1.json`…`node-6.json` have `"emails": []` (and some use a made-up
  cast like `{"CEO": "you"}`). These are placeholders — either fill them with real emails or
  delete them. Do not ship empty nodes.
- **Duplicate/near-duplicate nodes**: e.g. `Company_Retreat.json` vs `Company-Retreat.json`.
  Pick one, delete the other. Watch for other underscore-vs-hyphen twins.
- Inconsistent tiering and possibly answer keys that don't trace to a token.

## What "done" means

Every node must satisfy all of:

1. **Loads + lints clean.** `load_corpus` runs the linter (`sb/schema.py:lint`). No errors.
2. **Oracle is 100%.** The oracle solves the corpus from the answer keys alone; anything
   under 100% means a test is unwinnable and must be fixed.
3. **Grammar-correct.** Every date in an answer key traces to a `{token}` in the body
   (no free-typed dates). Answer keys are `ops` verbs (`create`/`move`/`cancel`) on named
   obligations — NOT the dead `expect`/`count`/`@HH:MM`/`duration` grammar.
4. **Tiered.** Each email carries an intended tier. Aim for roughly **30% T1 / 40% T2 /
   30% T3** across the whole corpus.

## The validation loop (offline, no API calls)

```bash
.venv/bin/python -m sb.scale --filler 120 --seed 42 --days 200 --dst build/scaled
```

This copies the corpus, lints on load, buries it in filler, and runs the oracle. The last
line must read:

```
oracle: N/N = 100% (must be 100% — corpus is valid at scale)
```

Iterate: fix a node → re-run → repeat until it lints clean AND prints 100%. Run the unit
tests too if you change anything structural: `.venv/bin/python -m pytest sb/tests -q`.

---

## Authoring discipline (the part that actually matters)

The single most important insight from prior runs: **date math does NOT discriminate
models.** A floor model already nails business-day chains and reschedules. The two levers
that separate strong from weak models are:

1. **Retrieval span** — how far back the needed fact was set (recent = easy; scrolled out of
   context, must `search_inbox` = hard). Span is mostly a *serving* knob (filler), partly
   orthogonal to tier.
2. **Task recognition / under-action** — the real failure mode is a model reading an implied
   obligation as a harmless FYI and doing nothing. This is where the discrimination lives.

So difficulty climbs along *those* axes, not "make the date arithmetic gnarlier."

- **T1** = self-contained, obvious action (or a plainly non-actionable FYI → `ops: []`),
  date from `serve` only, no dependencies.
- **T2** = one recent dependency, one anchor+offset or a reschedule; action is clear once you
  recall the one fact it points at.
- **T3** = stack **≥2** of: {fact buried far back, needs a second email's constraint
  (blackout/policy), action only implied, not stated}. Spend your best effort here. Author
  lots of the "implied obligation that looks like an FYI" pattern (Hardener C in TIER_LIST).
- **Anti-pattern:** a bait no-op email that stamps "no action needed" on itself — that traps
  nobody. Real ambiguity means a reasonable person could go either way until they think.
- **Keep it human-written.** No templated needles; the ambiguity must read like natural,
  slightly-under-specified email prose.

## `match`-keyword discipline (grading reliability — read carefully)

The grader finds the model's object by **substring-matching the model's *invented*
calendar/todo title against the obligation's `match` keywords** (case-insensitive, scoped to
the node, must be **exactly one** match). `match` defaults to `[name]`. A poorly chosen
keyword grades a correct action as wrong.

- Pick a `match` keyword the **email body itself leans on**, so any sane title echoes it.
- Prefer **one distinctive noun** over a multi-word set (ALL keywords must match).
- Name the obligation something a natural title would contain, then lean on the default.
- Within a node, give two same-kind obligations **distinct** keywords (the linter blocks a
  keyword set that fully catches another, e.g. `["review"]` vs `["client review"]`).

---

## Working rules

- Don't invent schema fields. If the grammar can't express something, note it in your final
  report rather than hacking around it.
- Make changes node-by-node and keep the oracle at 100% as you go — don't let it go red and
  pile up.
- The fake clock always starts **June 1, 2026** (day 0); it never reads the real date.
- When finished, summarize: which nodes you kept/merged/deleted, the final node + email
  counts, the tier distribution, and the final oracle line.
