# K — corpus authoring drift

Scope: free-typed dates, malformed / mis-rendering tokens, prose that contradicts its own
answer key, unused anchors, `kind` choices, `Innovation-comp` dominance.

**Method.** Every number below comes from an offline sweep of all 15 nodes / 167 emails,
not from reading one node. The corpus was loaded, the serve plan built
(`start_date=2026-06-01`, `seed=42`, default levers unless stated), every body rendered
through `resolver.render_body`, and every answer predicate resolved against its email's
own serve context. Each finding carries a one-liner that reproduces its headline number
from the repo root with `.venv/bin/python`. No corpus or code file was modified.

**Framing.** The corpus lints clean (`sb/schema.py:490`) and the oracle scores 167/167 at
seeds 42, 1, 99 and 2026. Every defect below is therefore invisible to both existing
checks. That is the finding behind the findings: lint validates *shape*, the oracle
validates *satisfiability*, and nothing validates that the rendered email a model reads
is consistent with the answer key it is graded against.

---

## K-1 The serve plan drifts past the dates the corpus narrates
Status: open
Severity: blocks-measurement
Cost to verify: free-offline

**What's wrong.** 18 of 167 emails (11%) are delivered on a day *after* a date stated in
their own rendered body, and 11 answer ops demand a calendar object dated before the email
that asks for it arrives. The cause is structural: 143 of 167 emails carry no `date`
dependency edge, so the scheduler has nothing forcing a date-bearing email to be served
before its date arrives, and `lint` check #4 only demands a `date` edge when an answer
references an *anchor* — a body `{dom:13,0m}` paired with an answer `eq: dom:13,0m`
references no anchor and is waved through.

**Evidence.**

```
.venv/bin/python -c "
from sb.schema import load_corpus
c=load_corpus('corpus')
print(len([e for e in c.emails.values() if not any(d.type=='date' for d in e.depends_on)]),'/',len(c.emails))"
→ 143 / 167
```

Blind spot sourced at `sb/schema.py:530-549` (check #4 keys off `_email_predicates` →
`_refs_in`, i.e. `@anchor` references only) and `sb/scheduler.py:76-110` (`has_date_edge`
→ `date_edge_ids` → `is_ready`; an email with no date edge is never deadlined and never
forced).

Answer ops dated before their email arrives, at the recorded lever (`daily_max=5`,
seed 42):

| email (`corpus/nodes/…`) | serve | op | resolves to | days in the past |
|---|---|---|---|---|
| `World_Cup_Cleat_Launch.json:347` press-briefing-and-embargo-2 | 2026-07-27 | move `Press Briefing` | 2026-07-13 | 14 |
| `Planning.json:12` rebrand-pitch | 2026-06-27 | create `vision_sync` | 2026-06-15 | 12 |
| `Sponsoring-Marathon.json:156` sponsorship-tiers | 2026-06-29 | create `approval of tier` | 2026-06-19 | 10 |
| `Company-Retreat.json:72` in-town-and-would-love-to-connect | 2026-06-28 | create `Athlete Meeting` | 2026-06-22 | 6 |
| `World_Cup_Cleat_Launch.json:272` reveal-event-date-and-venue-3 | 2026-07-23 | move `Reveal Rehearsal` | 2026-07-19 | 4 |
| `World_Cup_Cleat_Launch.json:272` reveal-event-date-and-venue-3 | 2026-07-23 | move `Reveal Event` | 2026-07-20 | 3 |
| `World_Cup_Cleat_Launch.json:589` board-sync-on-the-credit-issue | 2026-07-26 | create `Board Sync` | 2026-07-23 | 3 |
| `World_Cup_Cleat_Launch.json:13` project-design-kickoff | 2026-06-14 | create `Launch Window Decision` | 2026-06-12 | 2 |
| `World_Cup_Cleat_Launch.json:486` tooling-po-needs-approval | 2026-07-19 | create `Approve tooling PO` | 2026-07-17 | 2 |
| `Enterprise_Ai_Selection.json:100` final-review | 2026-06-06 | create `AI_Final_Meeting_Review` | 2026-06-05 | 1 |
| `Sponsoring-Marathon.json:25` launching-…-marathon-2 | 2026-06-26 | create `launchmeeting` | 2026-06-25 | 1 |

```
.venv/bin/python -c "
from datetime import date; from sb.schema import load_corpus
from sb.scheduler import build_plan; from sb import resolver; from sb.resolver import Context
c=load_corpus('corpus'); p=build_plan(c,start_date=date(2026,6,1),seed=42,n_days=200); n=0
for eid,sv in p.serve_date.items():
    ctx=Context(sv,p.anchors)
    for op in c.emails[eid].answer.ops:
        for k,v in (op.on or {}).items():
            if k in ('eq','by') and resolver._as_date(resolver.resolve(v,ctx))<sv: n+=1
print(n)"
→ 11
```

The prose degrades the same way. `World_Cup_Cleat_Launch.json:347` renders on 2026-07-27
as *"Since the reveal moved to the Monday, July 20th, 2026 … Moving the press briefing to
the Monday, July 13th, 2026 to keep it a week ahead of the new reveal date"* — both dates
already past, and the stated ordering (briefing a week ahead of the reveal) inverted.
`Enterprise_Ai_Selection.json:100` renders on Sat 2026-06-06 as *"lets meet Friday, June
5th, 2026 … I will email you closer to Friday"* and its answer is `eq: this:FRI` = the day
before delivery. `_ThisWD.eval` (`sb/resolver.py:145-148`) computes Monday-of-the-serve-week
plus the weekday offset, so `this:FRI` served on a Saturday or Sunday always resolves into
the past; 49 of 167 emails (29%) are delivered on a weekend.

Node-level correlation — the five nodes with **zero** date edges own 16 of the 18
past-narrating emails:

| node | emails | date edges | serve stretch | emails narrating a past date |
|---|---|---|---|---|
| `World_Cup_Cleat_Launch` | 22 | 0 | 47 d | 7 |
| `pizza-party` | 8 | 0 | 41 d | 5 |
| `Sponsoring-Marathon` | 12 | 0 | 51 d | 3 |
| `Enterprise_Ai_Selection` | 7 | 0 | 41 d | 1 |
| `shoe-product-launch-delays` | 6 | 0 | 43 d | 0 |
| `Company-Retreat` | 6 | 2 | 43 d | 1 |
| `Planning` | 6 | 0 | 34 d | 1 |
| *(remaining 8 nodes)* | 100 | 22 | — | 0 |

`corpus/nodes/pizza-party.json` is the extreme case. Rendered at seed 42 the eight emails
arrive in this order: the party is created for Mon Jun 8 (day 5); the *ordering deadline*
arrives Jun 19, eleven days **after** the party; the room-booking confirmation ("booked on
6/8") arrives Jun 26, eighteen days after; the reschedule request arrives Jul 5 asking to
move a party that already happened.

**Why it matters for benchmark validity.** An email asking a secretary to book something
last week has no correct answer. The model can satisfy the key only by doing something a
competent assistant would refuse. These emails are scored, they cannot be passed by correct
behaviour, and they are indistinguishable in the logs from genuine failure — `grader.py:177`
reports them as `"on the wrong day"`, the one bucket the brief (§2.3) treats as
"unambiguously measures temporal reasoning". At least some of the 18 opus / 18 sonnet
wrong-day failures are this, not reasoning.

**Options.**
1. Add a lint rule: every email whose *body* renders a date, or whose answer resolves to a
   date, must carry a `date` edge to whatever pins that date. Catches the class at
   authoring time; requires ~143 emails to gain edges, and over-constraining edges is what
   already makes 19% of seeds infeasible (K-2).
2. Add a scheduler-side invariant instead: `build_plan` fails if any email's resolved
   answer date precedes its serve date. Cheap, no corpus edits, but it converts the problem
   into "the corpus will not build" rather than fixing it, and would fail today at seed 42.
3. Re-author the affected emails to use anchor-relative expressions (`@retreat_date+2d`)
   rather than absolute month-day forms, so they stay coherent under any serve order.
   Highest fidelity, largest hand-authoring cost, and touches the corpus (see the
   never-change-grader-and-corpus-together rule in `CLAUDE.md`).
4. Accept and quarantine: tag the 18 emails, exclude them from the headline score, report
   them separately. Preserves comparability with the three recorded runs; shrinks n.

**Overlaps with:** K-2 (the count is lever-dependent), K-3 (some past-dating is also a
prose contradiction), G (past-dated ops surface as `"on the wrong day"` in the grader's
reason taxonomy), V (inflates the apparent temporal-reasoning failure rate).

**Open questions.** How many of the 18/18 opus/sonnet `"on the wrong day"` verdicts land on
these 11 ops? Answerable free-offline by joining `outputs/*.md` against the resolved plan,
but needs the per-email verdict parse that O-* proposes. Is the `this:WD`-into-the-past
behaviour (`resolver.py:145`) intended, or should `this:WD` clamp forward?

---

## K-2 The corpus is only coherent at `daily_max=21`; the recorded runs used 5
Status: open
Severity: blocks-measurement
Cost to verify: free-offline

**What's wrong.** The corpus's fixed content (free-typed prose, hard-coded `dom:` answers)
was evidently authored against `daily_max=21` — the lever pinned in `BENCHMARK_RESULTS.md`
§1 and used by the haiku run — but `outputs/opus.md` and `outputs/sonnet.md` ran at the
argparse default `daily_max=5` (brief §2.10 confirms the 57-day span). Because the whole
grammar is serve-relative by design (`sb/resolver.py:16`), changing only that lever moves
87 of 127 resolved answer dates and roughly triples the defect count. The two headline runs
were therefore scored against a corpus in a state it was never written for.

**Evidence.**

```
.venv/bin/python -c "
from datetime import date; from sb.schema import load_corpus
from sb.scheduler import build_plan, Levers; from sb import resolver; from sb.resolver import Context
c=load_corpus('corpus')
for dm in (5,8,13,21):
    p=build_plan(c,start_date=date(2026,6,1),seed=42,n_days=200,levers=Levers(daily_max=dm)); n=0
    for eid,sv in p.serve_date.items():
        ctx=Context(sv,p.anchors)
        for op in c.emails[eid].answer.ops:
            for k,v in (op.on or {}).items():
                if k in ('eq','by') and resolver._as_date(resolver.resolve(v,ctx))<sv: n+=1
    print(dm,p.n_days,n)"
```

| `daily_max` | days | prose contradicts its answer | past-dated ops | emails narrating a past date |
|---|---|---|---|---|
| 3 | **infeasible** | — | — | — |
| **5** (opus, sonnet) | 57 | 4 | **11** | **18** |
| 8 | 37 | 2 | 16 | 16 |
| 13 | 21 | 1 | 11 | 17 |
| **21** (documented, haiku) | 16 | 1 | **4** | **6** |
| 30 | 11 | 1 | 2 | 3 |

The single sharpest instance: `corpus/nodes/pizza-party.json:121` says *"Could we move the
pizza party to June 9"* and its answer is `eq: nth:2,TUE,0m`. At `daily_max=21` that
resolves to **2026-06-09** — the prose is correct. At `daily_max=5` the same email is served
2026-07-05 and the answer resolves to **2026-07-14**, 35 days from what the email asks for.
The free-typed date is not sloppiness; it is a correct date frozen against a lever the
recorded runs did not use.

87 of 127 resolved answer dates move between the two levers; 160 of 167 emails change serve
date. Largest shifts include `Enterprise_Ai_Selection.json:22` `Anthropic Zoom Call`
2026-07-21 → 2026-06-02 (49 d) and `World_Cup_Cleat_Launch.json:407` `Approve revised event
budget` 2026-07-22 → 2026-06-03 (49 d).

Separately, the corpus is not seed-robust:

```
.venv/bin/python -c "
from datetime import date; from sb.schema import load_corpus
from sb.scheduler import build_plan, InfeasibleSchedule
c=load_corpus('corpus'); bad=[]
for s in range(100):
    try: build_plan(c,start_date=date(2026,6,1),seed=s,n_days=200)
    except InfeasibleSchedule: bad.append(s)
print(len(bad),bad)"
→ 19 [5, 7, 13, 14, 17, 18, 23, 26, 27, 31, 32, 39, 41, 67, 69, 88, 92, 95, 97]
```

19 of 100 seeds strand emails permanently; `Company-Retreat`'s five date-edged emails
strand in 15 of the 19, and `daily_max=3` strands
`Rebrand-goes-company-wide.all-hands-to-get-everyone-aligned`. The mechanism is
`sb/scheduler.py:85-99`: a deadline is computed once and never recomputed, and
`is_ready` (`:108`) returns `False` forever once `today > dl`, so an email the random
filler fails to drain in time can never be served.

**Why it matters for benchmark validity.** The docstring at `sb/scheduler.py:15` promises
"Same (corpus, start_date, seed, levers) => identical plan, 100% reproducible". That is
true and beside the point: reproducible is not comparable. Two runs at different levers are
graded against materially different answer keys on the same 167 emails, so the haiku
59% and the opus 54% are not on the same scale. This is a live confound for the brief's
"central anomaly" (§2.1) and for §4.5's claim that phase 0 left the convergence untouched —
haiku ran the only lever at which the corpus's own prose is largely coherent. It does not
*prove* the anomaly is a corpus artefact (haiku also saw a different plan, different
context lengths, and 16 rather than 57 turns), but it removes "same benchmark, different
model" as a safe reading. It also means seed variance — the standard robustness check —
cannot be run without first fixing 19% of the seed space.

**Options.**
1. Pin `daily_max=21` as the canonical lever and re-run the roster there. Restores the
   corpus to the state it was authored in and is the cheapest path to a comparable set,
   but 16-day runs compress retrieval span further, which is the axis V cares about.
2. Pin `daily_max=5` and repair the corpus to be coherent at it (K-1 option 3). Keeps the
   longer, more realistic timeline; costs a full authoring pass.
3. Make the corpus lever-invariant: forbid free-typed dates and require every answer date
   to be anchor-relative, so no expression resolves against an arbitrary serve day.
   Eliminates the class permanently; the largest change and it removes a difficulty
   dimension (a model can no longer be tested on "the date in this email is stale").
4. Stamp lever + seed + corpus hash into every artifact and treat cross-lever comparison as
   invalid by convention. Zero corpus risk, no repair; the existing three runs stay
   incomparable.

**Overlaps with:** C (levers and corpus hash unstamped — this is *why* that matters), K-1
(same root cause, different lens), K-3, V (span is also lever-dependent: mean needle span
31.6 at seed 42 vs 43.5 at seed 2026).

**Open questions.** Was `daily_max=21` actually the authoring lever, or is the June-9
coincidence chance? Testable: resolve every free-typed date under a lever sweep and check
whether agreement peaks at 21 — I measured agreement at 5 and 21 only. Should
`build_plan` recompute deadlines each day rather than freezing them (`scheduler.py:91`)?
That is a scheduler change, not a corpus change, and would move the feasibility numbers.

---

## K-3 Prose states a date the answer key contradicts
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** Eight emails tell the model one date and grade it against another. Unlike
K-1 these are not artefacts of serve order alone — the body and the key disagree about
*which* date is meant, so a model that reads the email correctly is marked wrong. Two
distinct mechanisms: a free-typed month-day in prose that never traced to a token, and an
answer expression copy-pasted from an earlier email, which — being serve-relative — resolves
to a different day in its new home.

**Evidence.** Hand-verified list (seed 42, `daily_max=5`):

| # | email | prose says | answer resolves to | gap |
|---|---|---|---|---|
| 1 | `Company-Retreat.json:72` in-town-and-would-love-to-connect | "put the date of the meeting with the athlete on your calander for **August 22nd**" | `Athlete Meeting` `eq: dom:22,0m` = 2026-06-22 | 61 d |
| 2 | `pizza-party.json:119` client-demo-conflict | "move the pizza party to **June 9**"; "next Monday" = Jul 6 | `eq: nth:2,TUE,0m` = 2026-07-14 | 35 d / 8 d |
| 3 | `Day-of-execution_and_Aftermath.json:36` green-room | "green room an hour **before the livestream**" (`@launch_livestream` = 2026-08-10) | `eq: dom:10,+2m` = 2026-09-10 | 31 d |
| 4 | `Sponsoring-Marathon.json:221` pitch-deck-2 | "schedule a meeting for **July 2nd**" | `eq: dom:2,+1m` = 2026-08-02 | 31 d |
| 5 | `Company-Retreat.json:192` planning-call-and-forms | "join a call with me … on **June 21st**" | `eq: dom:21,0m` = 2026-07-21 (×2 ops) | 30 d |
| 6 | `Sponsoring-Marathon.json:184` approval-of-budget-tier | "meet **a week from today**" = 2026-07-14 | `eq: next:WED+2w` = 2026-07-22 | 8 d |
| 7 | `pizza-party.json:49` pizza-place-selection | "let me know which option **by Friday**" (Jun 26 or Jul 3) | `by: next:MON` → window \[Jun 28, Jun 29] | window excludes both |
| 8 | `pizza-party.json:94` pizza-order-deadline | "final order **by tomorrow**" = 2026-06-20 | `eq: this:FRI` = 2026-06-19 = the serve day | 1 d, wrong direction |

Row 3 is the cleanest diagnosis of the mechanism. `Day-of-execution_and_Aftermath.json:15`
emits `{!launch_livestream = dom:10,+2m}` and is served 2026-06-10, so the anchor holds
2026-08-10. `…json:36` copies the **same literal expression** `dom:10,+2m` into its answer
instead of writing `@launch_livestream`, is served 2026-07-22, and therefore resolves to
2026-09-10 — a "green room an hour before the livestream" scheduled a month after it. Two
answer expressions in the corpus duplicate an ancestor's emitted expression verbatim rather
than referencing the anchor:

```
.venv/bin/python -c "
from sb.schema import load_corpus, _email_predicates, _refs_in
c=load_corpus('corpus')
for eid,em in c.emails.items():
    emitted={x.strip():(n,a) for a in c.ancestors(eid) for n,x in c.emails[a].emits.items()}
    for e in (x.strip() for x in _email_predicates(em.answer)):
        if not _refs_in(e) and e in emitted: print(eid,e,'==@'+emitted[e][0])"
→ Company-Retreat.in-town-and-would-love-to-connect dom:22,0m ==@retreat_date
→ Day-of-execution_and_Aftermath.green-room-before-we-go-live dom:10,+2m ==@launch_livestream
```

Row 1 is internally three-way inconsistent: the body emits `{!athlete_visit = dom:22,1m}`
(= 2026-07-22), the prose says August 22nd, the answer says June 22nd, and a sibling email
(`Company-Retreat.json:119`) keys its own op to `@athlete_visit` = July 22.

**Detector honesty.** The automated relative-phrase detector flagged 9 emails; hand-checking
each against the full body left 4 true (rows 2, 6, 7, 8). The five false positives were all
the same shape — the phrase referred to a *different* event than the graded one
(`Planning.json:58` "I'm traveling **Thursday**" names the date being moved *from*;
`Pre-Launch.json:107` "customers walk into the new era **Monday** morning" names the day
after the reset; `Enterprise_Ai_Selection.json:65` and
`World_Cup_Cleat_Launch.json:182` both use "today" for context, not for the deadline).
Any lint rule built on this signal needs that discrimination or it will be ignored.

**Why it matters for benchmark validity.** These eight emails are unpassable by correct
reading. They land in the grader's `"on the wrong day"` bucket, which is the metric the
project cites as its temporal-reasoning signal, so they directly inflate the apparent
reasoning failure rate. Row 7 is worse than wrong: the `by` window is two days wide and
contains neither reading of "Friday", so every date the email actually names fails.

**Options.**
1. Lint rule: reject any body containing a month-name-plus-day or `M/D` outside a `{token}`.
   Purely syntactic, catches rows 1, 2, 4, 5 at authoring time, cannot see rows 3, 6, 7, 8.
2. Lint rule: reject an answer expression that is byte-identical to an ancestor's emitted
   expression (catches row 3 and the `retreat_date` case, 2 instances today). Narrow but
   exact and zero false positives.
3. Render-time cross-check: resolve the answer, then require at least one date in the
   rendered body to equal it (or the answer to reference an anchor whose value appears).
   Catches all eight; needs a whitelist for genuine misdirection emails and would fire on
   the 5 no-cue emails in K-6.
4. Hand-repair the eight and add no rule. Cheapest now; the class returns with the next
   authoring pass.

**Overlaps with:** K-1 (rows 1, 5 are also past-dated), K-2 (rows 2, 4, 5 flip with the
lever; rows 1, 3, 6, 7, 8 do not), K-5 (row 8's token also renders glued), G (the grader's
reason taxonomy cannot separate these from real errors).

**Open questions.** Rows 6 and 8 are 8-day and 1-day gaps — small enough that a tolerance
of `within:1d`/`within:7d` would mask them. Is masking acceptable, or does it hide a
different bug? Not answerable without the phase-1.5 hand-grade baseline.

---

## K-4 `Innovation-comp` supplies 29% of the emails and ~38% of every model's score
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** One of fifteen nodes contributes 48 of 167 emails but only 17 of 134
graded ops, has **zero** T3 emails, and is 31 of the corpus's 56 no-action emails. Because
a no-action email is passed by doing nothing, that node is the single largest source of
free points, and it drags every model's score toward the same number.

**Evidence.**

| node | emails | % | ops | no-action | T1 | T2 | T3 | create | move | cancel |
|---|---|---|---|---|---|---|---|---|---|---|
| `Innovation-comp` | **48** | **29%** | 17 | **31** | 18 | 30 | **0** | 16 | **0** | 1 |
| `World_Cup_Cleat_Launch` | 22 | 13% | 24 | 1 | 3 | 14 | 5 | 16 | 7 | 1 |
| `Sponsoring-Marathon` | 12 | 7% | 7 | 5 | 4 | 4 | 4 | 7 | 0 | 0 |
| *(12 more)* | 85 | 51% | 86 | 19 | 25 | 19 | 41 | 69 | 10 | 7 |
| **total** | 167 | | 134 | 56 | 50 | 67 | 50 | 108 | 17 | 9 |

Score share, parsed from the three committed logs (`✓ PASS` / `✗ FAIL` lines, 167 verdicts
each, totals matching the brief §2.1 exactly at 90/91/98):

| | opus | sonnet | haiku |
|---|---|---|---|
| `Innovation-comp` passes | 32/48 (67%) | 36/48 (75%) | 36/48 (75%) |
| …as a share of the total score | **36%** | **40%** | **37%** |
| no-action emails passed | 51/56 | 51/56 | 56/56 |
| …as a share of the total score | **57%** | **56%** | **57%** |
| action emails passed | 39/111 (35%) | 40/111 (36%) | 42/111 (38%) |

Strip the no-action emails and the three models score 35% / 36% / 38% — a 3-point spread on
the half of the benchmark that requires acting, versus the 5-point spread on the headline.
`Innovation-comp` also carries the least machinery of any node: 6 of its 48 emails contain a
date token, 0 emails have more than one dependency edge, and 0 of its 17 ops is a `move`.

Per-node pass rates also expose two nodes no model can do: `Partnership-with-deeptech-companies`
scores 2/10 for **all three** models, `World_Cup_Cleat_Launch` 7-8/22. A defect that all
three models hit identically is more likely corpus or grader than capability.

**Why it matters for benchmark validity.** 57% of every recorded score is "correctly did
nothing", and more than half of that comes from one node. The measured quantity is
therefore mostly a no-action-restraint test wearing a temporal-reasoning label, and the
Opus/Sonnet/Haiku convergence the brief calls the central anomaly (§2.1) is partly just
the shared ~51 free points. This is a distinct mechanism from the grader-collision one in
§4.1 and can be separated from it, since a no-action pass involves no title matching at all.

**Why it is not conclusive.** I have not established that `Innovation-comp`'s emails are
*bad*, only that they are numerous, easy, and homogeneous. A high no-action fraction is
defensible design (real inboxes are mostly noise); what is not defensible is that the
fraction is unstated and unreported. `sb/analyze.py` never reads `email.tier` (brief §2.9),
so no existing tool would show any of this.

**Options.**
1. Report score split by `has_ops` and by node as standard output. Zero corpus change,
   makes the composition visible, does not change any number.
2. Rebalance `Innovation-comp` down toward its ~13% op share by moving surplus no-action
   emails into `sb.scale`'s filler pool (`sb/scale.py:67-96`), where they already belong
   conceptually. Preserves total emails; changes the corpus and so every score.
3. Weight the score by op count rather than per-email binary, so a 4-op T3 email is not
   worth the same as an FYI. Changes the metric definition and breaks comparability with
   all three recorded runs.
4. Leave the composition and add T3 content to `Innovation-comp` so its 29% share buys
   proportionate difficulty. Largest authoring cost.

**Overlaps with:** V (binary per-email metric, tier data unread), G (no-action grading is
the one path that never touches `_title_hit`), A (`grader.py:186-195` no-action grading
depends on `email_id` attribution, so these free points are also the ones the sibling-stamp
hole could silently protect), C (composition is not stamped anywhere).

**Open questions.** Is `Innovation-comp` an intentional haystack node, or did it grow by
accident? Nothing in `docs/` says. Its 30 T2 tags on 31 no-action emails suggest the tiers
were applied without reference to whether the email requires an action at all.

---

## K-5 Rendered tokens fuse into prose, render at the wrong grain, or leak authoring notes
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** `resolver.human()` (`sb/resolver.py:368-378`) has exactly one output
shape — `"Monday, June 22nd, 2026"` — and authors wrote prose expecting other shapes or no
shape at all. The result is 9 emails where a full date is fused to an adjacent word, 8
places where a determiner precedes a full weekday-date, 2 naked `{serve}` tokens that
stamp today's date onto the end of a sentence, and one email that ships an unexpanded
authoring note verbatim.

**Evidence.**

```
.venv/bin/python -c "
import re; from datetime import date
from sb.schema import load_corpus; from sb.scheduler import build_plan
from sb import resolver; from sb.resolver import Context
c=load_corpus('corpus'); p=build_plan(c,start_date=date(2026,6,1),seed=42,n_days=200)
W='Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday'; n=0
for eid,sv in p.serve_date.items():
    r=resolver.render_body(c.emails[eid].body,Context(sv,p.anchors)).text
    for m in re.finditer(rf'(?:{W}), \w+ \d+\w+, \d{{4}}',r):
        pre=r[m.start()-1:m.start()]; post=r[m.end():m.end()+1]
        if (pre and not pre.isspace() and pre not in '(\"') or (post and not post.isspace() and post not in '.,;:!?)\"-'): n+=1
print(n)"
→ 9
```

Glued renders (rendered text, seed 42):

| email | rendered fragment |
|---|---|
| `Company-Retreat.json:14` | "We have chosen June 22nd**Monday, June 22nd, 2026** as the official start" |
| `pizza-party.json:12` | "Can't wait!!**Saturday, June 6th, 2026**" |
| `pizza-party.json:94` | "final order by tomorrow**Friday, June 19th, 2026** so they have time" |
| `Sponsoring-Marathon.json:13` | "Best, CEO**Saturday, June 20th, 2026**" |
| `Pre-Launch.json:14` | "get them done by**Monday, August 3rd, 2026**!" |
| `Company-Retreat.json:72` | "**Wednesday, July 22nd, 2026**Hey Man!" |
| `press-tour.json:110` | "The keynote moves to **Sunday, September 13th, 2026**[insert date → the 13th, +2 months, + add time 11:00 AM-12:00 PM]." |
| `World_Cup_Cleat_Launch.json:378` | "**Friday, July 17th, 2026**Hi Mark," |
| `Sponsoring-Marathon.json:209` | "We need this by**Thursday, July 30th, 2026** so that we have a week" |

`press-tour.json:110` is the worst of these: the bracketed note is authoring scaffolding
that was never removed, it reaches the model verbatim in the served email, and it states
the answer-key expression (`the 13th, +2 months`) in plain English next to the rendered
date. Grep confirms it is the only one: `grep -rn "insert date\|TODO\|TBD\|FIXME"
corpus/nodes/` returns one hit.

Grain mismatch — 8 occurrences across 5 emails where the author wrote `the {dom:N,…}`
expecting `"the 13th"` and got a full date: `World_Cup_Cleat_Launch.json:211, 239, 272 (×3),
314, 347 (×2)`. Rendered, `…json:272` reads *"won't arrive in time for the Monday, July
13th, 2026"*. `Partnership-with-deeptech-companies.json:110` has the same problem with an
Interval: `{!boston_trip_start = week_of:(dom:6,+1m)}` renders through
`resolver.human():371` as *"from **the week of August 3rd, 2026** till Sunday, August 9th,
2026"*, where the prose frame ("from X till Y") wants a single day.
`World_Cup_Cleat_Launch.json:314` renders "by next **next Tuesday, July 7th, 2026**".

Naked `{serve}` tokens (`pizza-party.json:14`, `Sponsoring-Marathon.json:15`) parse cleanly —
`_parse_base` accepts the bare literal `serve` at `sb/resolver.py:239` — so lint check #2
(`sb/schema.py:498-506`) passes them. Both sit at the very end of a body and stamp the
delivery date after the sign-off. Note the brief describes the `pizza-party` one as
"rendering a date mid-sentence"; it is at the end of the body, not mid-sentence.

**Why it matters for benchmark validity.** The rendered body is the entire input the model
under test receives. Fused text ("by tomorrowFriday, June 19th, 2026") is genuinely
ambiguous — it is not clear whether the deadline is tomorrow or the stated Friday — and in
that email the two differ, which makes it a K-3 contradiction as well. A benchmark that
scores comprehension of text it did not proofread is measuring its own typos. Separately,
`press-tour.json:110` leaks answer-key structure into the prompt, which is a scoring
integrity problem regardless of legibility.

**Options.**
1. Add a render-time lint: fail if a rendered date is adjacent to a word character, or if
   a body contains a `[…]` bracket note. Cheap, purely mechanical, catches all 10 above;
   would need to run per-serve-plan since rendering depends on the plan.
2. Give `human()` a short form (`"the 13th"`, `"June 22nd"`) selected by a token modifier
   (`{dom:13,0m|short}`). Fixes the grain class properly; changes the token grammar, which
   `sb/resolver.py:6` calls the keystone and which the webapp's generated TS types mirror
   (`sb/schema.py:29-31`).
3. Hand-fix the 12 affected emails (add a space, delete the two `{serve}` tokens, delete the
   bracket note, rephrase the 8 determiners). One authoring pass, no code change, no rule.
4. Do nothing about grain and only remove the leak in `press-tour.json:110`. Minimal;
   leaves the ambiguity in `pizza-party.json:94`, which is load-bearing for a graded op.

**Overlaps with:** K-3 (`pizza-party.json:94` is both), K-8 (same authoring-hygiene root),
V (a leaked answer expression undermines any retrieval claim about that email).

**Open questions.** Does a rendered date fused to a word actually change model behaviour, or
do models read through it? Testable free-offline only by inspection of the existing logs;
`press-tour.keynote-slot-swapped` and `pizza-party.pizza-order-deadline` verdicts across the
three runs would be the first place to look.

---

## K-6 Answers the prose does not pin
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** Five emails demand an exact calendar day while containing no temporal cue
of any kind — no token, no anchor reference, no weekday, no month, no "tomorrow". At the
other extreme, twelve `by:` answers accept a window averaging 19.5 days (max 64), and two
`eq:` answers resolve against an Interval anchor, which `grader.py:78-79` silently
reinterprets as "anywhere in that week". Both ends make the answer key say something other
than what it appears to say.

**Evidence.** Exact-day answers with no cue in the body:

| email | tier | answer | body (opening) |
|---|---|---|---|
| `Innovation-comp.json:905` found-a-typo-on-the-trophy | T1 | `create` event `eq: serve+1d` | "…just need a quick approval on the corrected text before we reorder." |
| `Innovation-comp.json:638` trophy-design-quick-look | T1 | `create` todo `eq: serve+3d` | "…let me know if anything feels off, otherwise we'll go ahead." |
| `Innovation-comp.json:411` need-your-sign-off-on-prize-amounts | T1 | `create` todo `eq: serve+2d` | "Just need a yes or no, shouldn't take long." |
| `Sponsoring-Marathon.json:184` approval-of-budget-tier | T2 | `create` event `eq: next:WED+2w` | "meet a week from today" (→ also K-3 row 6) |
| `Day-of-execution_and_Aftermath.json:36` green-room | T2 | `create` event `eq: dom:10,+2m` | "an hour before the livestream" (→ also K-3 row 3) |

The three `Innovation-comp` rows are `serve+1d` / `serve+2d` / `serve+3d` on bodies whose
only urgency cue is "quick" or "shouldn't take long". `tolerance` is `exact_day` on all
three (`sb/schema.py:151`, the default), so an assistant who files the task for the same
afternoon fails. Nothing distinguishes +1 from +3 in the text.

`by:` window widths (days the grader accepts, `grader.py:96`:
`ctx.serve <= obj.when.date() <= by_date`):

| window | email | op |
|---|---|---|
| 64 d | `Day-of-execution_and_Aftermath` metrics-readout | `by: @launch_livestream+3d` |
| 44 d | `Rebrand-goes-company-wide` all-hands | `by: @acme_reveal-3d` |
| 39 d | `Pre-Launch.json:14` launch-day-locked | `by: dom:3,+2m` |
| 30 d | `Pre-Launch` design-locked | `by: @launch_day-1w` |
| 20 d | `Company-Retreat.json:119` athelete-visit | `by: @athlete_visit` |
| … | *(7 more, 2–11 d)* | |

n=12, mean 19.5 d, max 64 d. A 64-day window on a 57-day run is satisfied by essentially any
date the model picks, so that op tests nothing but whether an object with the right keyword
exists.

Interval-as-`eq`: `Partnership-with-deeptech-companies.json:110` keys `Boston Trip Start` and
`WHOOP HQ Visit` to `eq: @boston_trip_start`, where the anchor holds
`week_of:(dom:6,+1m)` = Interval(2026-08-03 … 2026-08-09). `_matches_value`
(`sb/grader.py:78-79`) tests `expected.contains(...)` for an Interval regardless of
tolerance, so both `eq` ops silently accept any of seven days while the prose says
"the first day".

**Why it matters for benchmark validity.** The corpus is simultaneously too strict and too
loose, and neither is visible from the answer key's surface form. The five no-cue emails are
unpassable except by luck (`serve+2d` is one guess among many); the wide `by:` windows are
unfailable. Both distort the score away from what the benchmark claims to measure, in
opposite directions, which is worse than a consistent bias because it does not cancel.

**Options.**
1. Lint rule: an `eq` answer must be traceable to something in the body — a token, an
   anchor reference, or a date phrase. Catches the five; needs a definition of "date
   phrase" and will false-positive on deliberate-ambiguity emails.
2. Cap `by:` windows at a stated maximum (e.g. 14 days) and lint the rest. Simple; some
   long-horizon obligations are legitimately open-ended.
3. Forbid `eq` on an expression that can resolve to an Interval, or make `_matches_value`
   raise instead of silently widening. Small and exact; a grader change, so it must not
   ship in the same commit as a corpus change (`CLAUDE.md`).
4. Convert the five no-cue answers to `within:Nd` tolerance rather than repairing the prose.
   Cheapest; concedes that the email does not specify a day, which arguably it should.

**Overlaps with:** G (tolerance and predicate semantics are the grader's contract; option 3
is a grader change), K-3 (two rows appear in both), V (a 64-day window contributes no
temporal signal).

**Open questions.** Are the three `Innovation-comp` `serve+Nd` answers meant as a
"reasonable-turnaround" test with implied tolerance? If so `exact_day` is the wrong default
for them, and the same question applies to every `serve+Nd` op in the corpus — I did not
count those.

---

## K-7 `kind` and calendar-plausibility choices
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** Four `create` ops declare a `kind` that contradicts what the obligation
plainly is, most visibly a *party* keyed as `kind: todo`. Because `grader.py:163` filters
the candidate pool by `op.kind` before matching titles, a model that creates the right
object of the other kind scores `"(nothing matching created)"` — identical in the log to
having done nothing. Separately, 14 of 112 `eq` ops land on a Saturday or Sunday.

**Evidence.**

| email | obligation | `match` | declared | reads as |
|---|---|---|---|---|
| `pizza-party.json:12` end-of-year-pizza-party | `Team_pizza_party` | `["Team_pizza_party"]` | `todo` | an event (a party, "having you there") |
| `Enterprise_Ai_Selection.json:100` final-review | `AI_Final_Meeting_Review` | `["ai"]` | `todo` | an event ("lets meet {this:FRI}") |
| `Company-Retreat.json:119` athelete-visit | `Contact People Added To List` | `["people"]` | `event` | a to-do ("make sure to contact the people") |
| `Innovation-comp.json:905` found-a-typo | `approve_trophy_correction` | `["correction"]` | `event` | a to-do ("just need a quick approval") |

Detector: obligation name + `match` keywords matched against an event-word list
(party/meeting/call/…) and a to-do-verb list (send/approve/contact/…); 4 of 108 `create`
ops flagged, each hand-confirmed against its body above. The detector is deliberately
conservative and will have missed ambiguous cases.

`pizza-party.json:12` is the sharpest: the same node's `client-demo-conflict` (`:119`)
issues a `move` on that obligation, and `_wire_obligations` (`sb/schema.py:393`) makes the
`move` inherit `kind: todo`, so a model that reasonably books the party as a calendar event
fails **both** the create and the move — two of the node's four graded emails — with
`"no to-do titled like Team_pizza_party was created/moved"`.

Weekend placement, 14 of 112 `eq` ops, including `press-tour.json:110` `expo_keynote`
(a public keynote) on Sun 2026-09-13, `World_Cup_Cleat_Launch.json:272` `Reveal Rehearsal`
on Sun 2026-07-19 the day before a Monday reveal, `World_Cup_Cleat_Launch.json:431`
`Design Freeze Sign Off` on Sun 2026-07-19, and `Enterprise_Ai_Selection.json:151`
`AI_Sign_Off` on `eq: next:SUN` — the only one where the weekend is explicit in the
expression. The rest fall on weekends because `dom:N`, `serve+Nd` and anchor arithmetic have
no business-day awareness, even though `add_business_days` exists (`sb/resolver.py:80`) and
`+Nbd` is in the grammar.

29% of emails are themselves delivered on a Saturday or Sunday, which is what puts
`this:WD` answers into the past (K-1).

**Why it matters for benchmark validity.** A kind mismatch is one of the four causes the
brief says the log cannot distinguish (§2.3, §4.3), and here are four cases where the corpus
itself is the cause rather than the model. Weekend-dated obligations penalise a model
applying ordinary business-calendar priors — the priors a real assistant should have — and
that penalty falls on the same `"on the wrong day"` bucket used as the temporal-reasoning
signal.

**Options.**
1. Re-key the four ops to the kind the prose implies. Four-line corpus change; changes
   scores, so it must not land with a grader change.
2. Make the grader kind-tolerant (search both pools, penalise only date/count). Removes the
   whole ambiguity class including cases the corpus does not cause; a substantial change to
   the grading contract and G's territory.
3. Add a lint heuristic warning on event-words keyed `todo` and vice versa. Advisory only;
   the word lists are judgement calls and will annoy authors.
4. Add a business-day lint (warn when an `eq` op lands on a weekend unless the expression
   names one) and/or route the affected answers through `+Nbd`. Independent of the kind
   question; 14 instances.

**Overlaps with:** G (the kind filter at `grader.py:163` is the mechanism; option 2 is a
grader change), K-1 (weekend serve days drive `this:WD` into the past), A (a kind mismatch
and a mis-stamped `email_id` produce the same log line).

**Open questions.** Do the three recorded runs actually fail the `Team_pizza_party` create
as a kind mismatch? All three score pizza-party 4/8; the per-email verdicts would say which
four, and that is free-offline once the log parse from O-* exists.

---

## K-8 Anchor, name and metadata hygiene
Status: open
Severity: slows-work
Cost to verify: free-offline

**What's wrong.** 11 of 47 emitted anchors (23%) are referenced by nothing, including a
pair that differs only in capitalisation and resolves to two different dates for the same
event. Nine obligation names carry leading or trailing whitespace, 13 subject lines are
duplicated across the corpus, and 3 emails have an empty `from`.

**Evidence.**

```
.venv/bin/python -c "
from sb.schema import load_corpus
c=load_corpus('corpus'); refs=set()
for e in c.emails.values(): refs|=e.anchor_refs
u=[(n,s) for n,s in c.emission_map.items() if n not in refs]
print(len(u),'/',len(c.emission_map)); [print(' ',n,'<-',s) for n,s in sorted(u,key=lambda x:x[1])]"
→ 11 / 47
```

Unused: `@Anthropic_Meet`, `@Review_Meeting` (`Enterprise_Ai_Selection`); `@board_signoff`,
`@vision_sync`, `@rebrand_reveal` (`Planning`); `@launchmeeting`, `@approval_of_tier_budget`
(`Sponsoring-Marathon`); `@ordering_date`, `@pizza_decision`, `@New_party_date`,
`@new_party_date` (`pizza-party`). The brief's §2.10 says pizza-party has "three emitted
anchors nothing references"; it has **four**.

`corpus/nodes/pizza-party.json:149` emits `{!new_party_date = dom:9,0m}` → 2026-07-09 and
`:170` emits `{!New_party_date = nth:2,TUE,0m}` → 2026-07-14. Two anchors for the same party
date, distinct only by case, holding dates five days apart, while the graded answer at
`:121` is 2026-07-14. Rendered, `:149` reads *"Thursday, July 9th, 2026 works for me, thanks
for the heads up"* and `:170` reads *"isn't available on Tuesday, July 14th, 2026"*. Together
with the free-typed "June 9" at `:121`, that node states **three** different party dates.
`_build_emission_map` (`sb/schema.py:451-460`) raises only on exact-name collisions, so the
case pair is accepted.

An unused emit is not free: `emits` feed the global anchor table
(`sb/scheduler.py:138-139`) and `_build_emission_map` raises on any duplicate name across
the whole corpus, so every unused anchor is a name reserved corpus-wide for nothing. Note
`{dom:9,0m}` renders identically to `{!new_party_date = dom:9,0m}` — the emit syntax buys
nothing when the name is never referenced.

Obligation names with edge whitespace — 9 distinct names over 10 ops
(`Innovation-comp` 3, `Marketing-campaign-new-product-delay` 4, `Enterprise_Ai_Selection` 2):
`Innovation-comp.json:649` `" review_trophy_design"`, `:740` `" thank_judges"`,
`:864` `" retention_conversation"`, `Enterprise_Ai_Selection.json:22`
`"Anthropic Zoom Call "` and `"OpenAI In Person Walk Through "`, and four in
`Marketing-campaign-new-product-delay.json` including
`"Serena William marketing campaign "` (used by both a `create` and a `move`).
Because `match` defaults to
`[name]` (`sb/schema.py:155`), a trailing-space name becomes a **match keyword containing a
trailing space** — `_title_hit` (`grader.py:69-70`) then requires that space to appear in
the haystack, which it does for any multi-word title but not for a title ending on that
word.

13 duplicated subject lines, e.g. `"Reveal event date and venue"` ×3 and
`"Manufacturing kickoff"` ×3 in `World_Cup_Cleat_Launch`, `"Pitch Deck"` ×2 in
`Sponsoring-Marathon`. `list_new_emails` returns ids and subjects, so a model choosing what
to `get_email` sees three indistinguishable rows. Empty `from`:
`Innovation-comp.json:893`, `pizza-party.json:49`,
`Sponsoring-Marathon.json:221`. Ten subject lines carry edge whitespace.

**Why it matters for benchmark validity.** Individually cosmetic; collectively this is the
authoring-quality signal that says the other findings are not isolated slips. The case-pair
anchor and the whitespace-in-`match` cases are the two that can actually change a verdict.
Duplicated subjects degrade the retrieval task the benchmark claims to test.

**Options.**
1. Lint: reject unused emitted anchors, anchor names differing only by case, and names with
   edge whitespace. Purely mechanical, no judgement, would fail the corpus today at 11 + 1 + 9
   sites.
2. Strip whitespace at parse time in `_parse_op` (`sb/schema.py:130-156`) and warn. Fixes the
   `match`-keyword consequence without touching the corpus; hides the authoring error.
3. Lint duplicate subjects within a node, or auto-suffix them ("Reveal event date and venue
   (update 2)"). Improves the retrieval surface; changes what the model reads, so it changes
   scores.
4. Hand-clean and add no rules. Fast; the class returns.

**Overlaps with:** G (`match` defaults and keyword quality — the trailing-space keywords and
the short keywords `["ai"]`, `["end"]`, `["fbs"]`, `["who"]` are G's `_title_hit` problem,
listed here only as authoring provenance), K-5.

**Open questions.** Does any of the 9 edge-whitespace names actually change a verdict in the
three recorded runs? Needs the per-email verdict parse.

---

## Reproduction notes for the synthesizer and verifier

**§2 claims I reproduced.** §2.10 in full: 15 nodes / 167 emails / 134 graded ops; tiers
50 T1 / 67 T2 / 50 T3; `Innovation-comp` 48 emails (29%), 17 ops, zero T3;
`sb.scale --filler 0 --seed 42 --days 200` → 167 emails, 57 days, `oracle: 167/167 = 100%`;
span mean 31.6 / max 83 / n=24. §2.1's three scores (90 / 91 / 98 of 167) reproduce exactly
by parsing the `✓ PASS` / `✗ FAIL` lines in `outputs/opus.md`, `outputs/sonnet.md`,
`past/claude-haiku-4-5.md`.

**§2 claims I could not reproduce as stated.**

- §2.10 says pizza-party has "three emitted anchors nothing references". It has **four**
  (`@ordering_date`, `@pizza_decision`, `@new_party_date`, `@New_party_date`) — see K-8.
- §2.10 says `by tomorrow{!ordering_date = this:FRI}` renders as "tomorrowFriday, June 12th".
  The form is right; at the default config it renders **June 19th, 2026**, and the email is
  served *on* Friday June 19. June 12 corresponds to some other seed/lever, which means the
  brief's rendering was measured under a configuration it does not state.
- §2.10 calls the `{serve }` token "a stray token rendering a date mid-sentence". It is a
  valid token (`resolver.py:239` accepts bare `serve`) sitting at the **end** of the body,
  and there are **two** of them in the corpus, not one (`pizza-party.json:14`,
  `Sponsoring-Marathon.json:15`).
- §2.10's framing that the free-typed dates "do not trace to a token" is right about the
  syntax but wrong about the cause. `pizza-party.json:121`'s "June 9" is exactly what
  `nth:2,TUE,0m` resolves to at `daily_max=21`. These are correct dates frozen against the
  documented lever, not invented ones (K-2).

**§4 items I believe are now partly established.**

- **§4.5** ("phase 0 did not fix the convergence — that is the grader"). Not only the
  grader. Two corpus-side mechanisms are now measured: 57% of every recorded score is
  no-action passes (K-4), and the corpus's date coherence is a function of `daily_max`, with
  haiku the only run at the coherent lever (K-2). Neither proves the anomaly is a corpus
  artefact; both mean "same benchmark, different model" is no longer a safe reading of the
  three runs.
- **§4.3** (the split of the 52 "nothing matching created" cases). Still unmeasured, but the
  kind-mismatch arm now has four named corpus-caused candidates (K-7) rather than being
  purely hypothetical.

**§4 items I did not touch.** §4.1, §4.2, §4.4.

**New, outside the brief.** 19 of 100 seeds make the corpus unschedulable
(`InfeasibleSchedule`), as does `daily_max=3`; `Company-Retreat` strands in 15 of the 19.
Any plan to check score variance across seeds has to fix that first. Mechanism at
`sb/scheduler.py:85-99` (deadlines computed once, never recomputed) plus over-constrained
`date` edges in one node.
