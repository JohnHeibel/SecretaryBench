# VERIFY — red-team of the phase-2 identity contract, iteration 3 (`grade_v4`)

Adversarial pass over the **proposed, not-yet-implemented** iteration 3 of the grader
contract (`docs/grader-contract.md`, prototype `docs/_repair/phase2_guards.py::grade_v4`).
**This file reports; it changes nothing.** No edit to `sb/`, `corpus/`, `captures/`,
`webapp/`, the prototype, or any other doc. No git operation. No live model run.
`scripts/recover_corpus.py` and `scripts/fix_match.py` were neither read nor executed.

**Method.** I wrote my own scoring harness from the prose of `docs/grader-contract.md`
rather than calling `grade_v4`, so a bug in the prototype's measurement code could not
hide. It re-implements tokenisation, the stop list, stemming, keyword extraction,
overlap, the two-phase exclusive assignment, the kind-first tie-break, turn membership
by value, the volume brake, the stale-survivor rule, the cancel rule, the date
predicate, node/turn pool construction, the day loop and the synthetic agents. It
borrows from `sb/` only `resolver.resolve`/`Interval` (date arithmetic, identical under
every contract), `schema.load_corpus`, `scheduler.build_plan`, and `oracle._target`/
`_as_dt` for the synthetic agents' dates. I then ran the prototype's own
`report([...])` and compared. Scratch code lives outside the repo at
`/private/tmp/claude-501/-Users-jamesoc-dev-SecretaryBench/cf81dddf-4a9b-49d4-996d-ea65d999c189/scratchpad/verify3/`
— `ref3.py` (contract + worlds), `eng.py` (engine path), `adv.py` (new adversaries),
`cases.py` (constructed cases). Every command below is prefixed
`PYTHONPATH=/Users/jamesoc/dev/SecretaryBench:<scratchdir> .venv/bin/python`.

---

## Bottom line

**No — not safe to implement in `sb/grader.py` as specified.** Not because the design is
wrong in outline; the identity rule, the cancel rule and the diagnosis work are real
improvements and I confirmed almost every number in the document to the digit. It is not
safe because **two of its load-bearing safety properties do not hold**, and both are
stated in the document as facts.

1. **The volume brake (rule 3) is defeated by attribution laundering, on the only code
   path that scores real runs.** A model that does the work correctly and then leaves
   five duplicate copies of every object it creates — stamping each copy with a *sibling
   email's id from the same node* — scores **167/167** under iteration 3. Its store holds
   **509 objects where the perfect agent holds 99**. The shipped grader scores the same
   agent **72/167**, eight points above the do-nothing floor. `sb/live/runner.py:109`
   tells the model *"Never leave duplicates"*; iteration 3 makes that instruction
   unenforceable again. The document asserts the opposite in three places (rule 2 of
   `docs/grader-contract.md`, and both the 2026-08-24 and 2026-08-29 changelog entries):
   *"cannot be defeated by a mis-stamped `email_id` (register A-5)"*. It can, because on
   `sb/live/runner.py:611` and `sb/regrade.py:68` the turn delta **is** the email_id
   split. The same state scores **64** on the engine path and **167** on the runner path
   — a 103-point disagreement between the two grading paths, the exact defect iteration 3
   claims to have closed (it closed a 1-point instance of it and opened a 103-point one).

2. **The assignment is order-dependent, so `oracle_engine = 167` and `oracle_name = 167`
   are order-luck, not properties of the contract.** Shuffling the pool order of the
   captured state (a permutation that changes no object and no date) moves `oracle_name`
   to **165 / 166 / 167** — 167 in only 19 of 60 shuffles. The shipped grader returns 158
   in all 20 shuffles. A *scheduling-perfect* agent that handles one email's two `move`
   ops in the opposite order scores **165 through `sb.engine.run`**, so
   `sb/scale.py:127`'s mandatory `must be 100%` gate is decided by an artifact of the
   model's edit sequence. That also breaks the invariant `sb/grader.py:5-6` states in
   plain English: *"a reschedule done as update-in-place and one done as
   delete-then-recreate score identically."* Under iteration 3 they do not.

Both are fixable, cheaply, and both fixes are measured in §12: scoping the brake to the
**day** rather than the email_id-filtered turn takes `launder` from 167 to 77 at **zero
cost on every guard column**, and a **specificity-first** tie-break makes `oracle_name`
167 and `dupmove` 152 in 30 of 30 pool orderings at a cost of 2 points on `oracle_subject`.
Neither defect was visible to the guard set, for the same structural reason iteration 1
failed: **the guards vary the agent, never the bookkeeping.** Every world in
`phase2_guards.WORLDS` stamps `email_id` correctly and presents objects in creation order.

Also **WRONG, and cited as the reason for a design choice:** *"Overlap-first ordering was
measured and rejected: it fails a model that creates both an event and a to-do for one
obligation, which the shipped grader passes."* Measured: that model scores **166 under
overlap-first, 166 under kind-first, 166 under a kind-filtered pool, and 153 under the
shipped grader** — the shipped grader does not pass it, and overlap-first does not fail
it. The ordering that was rejected is the one that puts the `wrongkind` adversary at
exactly the null floor (64 vs 69).

| | shipped | **iter. 3** | null floor |
|---|---|---|---|
| real (certified 1c capture) | 97 | **114** | 64 |
| `oracle_name` (perfect agent) | 158 | **167** | 64 |
| `oracle_name`, pool order shuffled ×60 | 158 (all) | **165 / 166 / 167** | — |
| perfect agent, two moves handled in the other order, via `sb.engine.run` | 158 | **165** | — |
| **NEW `launder` — right answer + 5 copies of every created object, copies stamped with a node sibling's id** | **72** | **167** ✗ | 64 |
| the same agent scored through `sb/engine.py:147` | 64 | **64** | 64 |
| **NEW `launder_all` — as above, duplicating on moves too** | **64** | **152** ✗ | 64 |
| **NEW `dupmove_retitle` — double-books every move, drops one word from the stale copy's title** | 146 | **163** ✗ (guard says ≤ 152) | 64 |
| **NEW `bothkinds` — creates an event *and* a to-do for every obligation** | 153 | **166** | 64 |
| **NEW `nocancel` — perfect, but never deletes anything** | 150 | 159 | 64 |

Rows marked ✗ are guard failures. `launder` meets the brief's regression criterion twice
over: it scores 167 without doing the work asked, and it beats the null floor under
iteration 3 (+103) while barely beating it under the shipped grader (+8).

---

## 1. Reproducing the measured table — CONFIRMED, every cell

`ref3.py` (my implementation) against `phase2_guards.report([...])` (the prototype).
Both were run; the two agree with each other and with `docs/grader-contract.md` lines
159-170 on **every cell of every row**.

| contract | real | oracle_engine | oracle_name | oracle_subject | oracle_inflect | null | dup5 | shot7 | shot45 | shot90 | wrongkind | dupmove |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| shipped — doc / prototype / **mine** | 97 | 167 | 158 | 95 | 142 | 64 | 64 | 64 | 64 | 64 | 65 | 148 |
| shipped + name-titled oracle | 97 | 158 | 158 | 95 | 142 | 64 | 64 | 64 | 64 | 64 | 65 | 148 |
| iteration 2 (`grade_v3`) | 114 | 166 | 167 | 138 | 159 | 64 | 64 | 64 | 64 | 64 | 69 | 167 |
| **iteration 3 — doc / prototype / mine** | **114** | **167** | **167** | **137** | **159** | **64** | **64** | **64** | **64** | **64** | **69** | **152** |
| iter. 3, cancel any-word | 112 | 164 | 164 | 134 | 157 | 64 | 61 | 61 | 61 | 61 | 66 | 149 |
| iter. 3, cancel ≥ 0.5 | 112 | 166 | 166 | 135 | 158 | 64 | 63 | 63 | 63 | 63 | 68 | 151 |
| iter. 3, no stale rule | 114 | 167 | 167 | 138 | 159 | 64 | 64 | 64 | 64 | 64 | 69 | 167 |
| iter. 3, kind-filtered pool | 114 | 167 | 167 | 137 | 159 | 64 | 64 | 64 | 64 | 64 | 69 | 152 |
| iter. 3, overlap before kind | 114 | 167 | 167 | 137 | 159 | 64 | 64 | 64 | 64 | 64 | 64 | 152 |
| iter. 3, title-only brake | 114 | 167 | 167 | 137 | 159 | 64 | 64 | 64 | 64 | 64 | 69 | 152 |

Engine-path numbers, independently (`eng.py`): shipped grader + `sb/oracle.py` as shipped
**167**; shipped grader + name-titled oracle **158**; iteration 3 + `sb/oracle.py` as
shipped **163**; iteration 3 + name-titled oracle **167**; cancel ≥ 0.5 **166**; cancel
any-word **164**. All four numbers the document quotes for blocker 3 reproduce exactly.

**Op-level failure profile — CONFIRMED, every cell** (n = 190 = 134 ops + 56 no-action
emails):

| bucket | shipped (doc / mine) | iter. 3 (doc / mine) |
|---|---|---|
| ok | 56 / 56 | 76 / 76 |
| not found | 50 / 50 | 18 / 18 |
| wrong kind | — | 8 / 8 |
| wrong day | 14 / 14 | 29 / 29 |
| count: too many | 11 / 11 | 0 / 0 |
| stale copy on move | — | 1 / 1 |
| cancel residue | 3 / 3 | 2 / 2 |
| over-acted | 2 / 2 | 2 / 2 |
| **failing ops** | **80 / 80** | **60 / 60** |

One presentational nit: the `ok` row silently excludes the 54 passing no-action details
(my raw counts are 110 → 130 of 190; 110 − 54 = 56, 130 − 54 = 76). It is internally
consistent, but a reader adding the column gets 136, not 190.

**"Iteration 3 changes no email verdict against iteration 2 (0 flips)" — CONFIRMED.**
Running `grade_v3` and `grade_v4` side by side over the capture: both 114, and the
passing sets are identical (0 emails differ).

**"`dupmove` = 152 = 167 − 15 exactly" — CONFIRMED and the guard is sensitive, not
coincidental.** Double-booking `k` of the 15 move emails gives iteration 3
`k=0 → 167, k=3 → 164, k=8 → 159, k=15 → 152`: exactly one point per double-booked move.
The shipped grader over the same sweep gives `158 / 156 / 151 / 148` — it misses 5 of the
15.

**One table cell in the document is WRONG.** `docs/grader-contract.md:161` marks the
shipped row's `dupmove` **148 ✓** and says the shipped contract "fails `oracle_name`"
(singular). The prototype's own guard checker disagrees, in the run I made from the
document's own command:

```
  shipped                    FAIL -> oracle_name 158, dupmove 148
```

148 > 158 − 15 = 143. The shipped grader fails the `dupmove` guard too, and the fact it
misses a third of the double-bookings is worth stating, not hiding behind a tick.

---

## 2. The volume brake is defeated by attribution laundering — UNSAFE

### 2.1 The exploit

Rule 3 fails a create/move when *"another same-kind object created **in the same turn**
matches it at least as well and was claimed by nobody."* The document justifies scoping to
the turn rather than to `email_id` with:

> Turn membership exists on every path and cannot be defeated by a mis-stamped
> `email_id` (register A-5). — `docs/grader-contract.md`, rule 2

That is false on the path that scores every live run and every offline re-grade:

```python
sb/live/runner.py:611    eid_new = {i for i in day_new if by_eid.get(i) == eid}
sb/live/runner.py:612-614  grade_email(..., _node_state(...), _turn_delta(corpus, state, eid_new))
sb/regrade.py:68         eid_new = {i for i in day_new if by_eid.get(i) == eid}
```

Turn membership on the runner and regrade paths **is** the `email_id` split. An object
stamped with any other id is not in `turn`, so `is_fresh` is false, so the brake never
sees it. It remains fully visible in the node pool (`sb/live/runner.py:139-148` keeps any
object whose stamp resolves to an email in this node), i.e. it is on the user's calendar
and it is counted against every other rule — just not against the duplicate rule.

The adversary (`adv.py::build(copies=5, launder='any', creates_only=True)`) is a
scheduling-perfect agent that, after creating the correct object for an obligation,
creates five identical copies stamped with another email id from the same node. It does
not duplicate on moves (see §2.3).

```
world                      final store   shipped   iter. 3   null
perfect (control)          99 objects       158      167       64
launder ×5 (creates only)  509 objects       72      167       64
launder ×1 (creates only)  ...               72      167       64
launder ×20 (creates only) ...               72      167       64
```

**167/167 with five times the calendar.** The count of copies is irrelevant; one is
enough, twenty is free.

### 2.2 The two grading paths disagree by 103 points on the same state

The identical agent, expressed as an `sb.engine` model and run through
`sb/engine.py:133-150` (where `turn = store.delta_since(before)` is the whole per-email
delta, `sb/engine.py:145`):

```
ENGINE path  (turn = full per-email delta):
   shipped grader : 64
   iteration 3    : 64
```

versus 167 on the runner path. Iteration 3's stated achievement over iteration 2 was
that *"the two grading paths disagreed on the same state"* and that turn membership fixes
it. It fixed a 1-email instance (`test_e2e`'s `beta.shift`) and introduced a 103-email
one. The `sb.scale` oracle gate runs on the engine path and would never see this; every
number in `docs/grader-contract.md`'s `real` column runs on the runner path and would.

### 2.3 Why the numbers are what they are

Laundering defeats the brake completely: under `launder_all` (duplicating on moves too,
`adv.build(copies=5, launder='any')`) the op-level profile is
`{'stale copy on move': 17}` and **zero** `over-created`. The only thing still catching
anything is rule 4, and only because the laundered copy created *this* turn is, by the
same email_id filter, classified as *inherited* rather than *fresh* — the brake's bug
accidentally feeding the stale rule. Not duplicating on moves removes even that, and the
score is 167.

The register already records the mechanism (A-5, A-3 route 1, A-4's blind spot: an id
belonging to a different email in *today's* batch produces no store warning at all). What
is new in iteration 3 is that **one mis-stamp now buys two rules**: the no-action check
(pre-existing) and the volume brake (new). A-5 must not be recorded as unaffected by this
contract; it is now the contract's largest gaming surface.

---

## 3. The assignment is order-dependent — UNSAFE

### 3.1 `oracle_name` = 167 does not survive a permutation

Shuffling the order of records inside each captured day's `state` (no object changed, no
date changed — only the list order the pool is built from):

```
real,           40 shuffles -> {114}                       canonical 114
oracle_name,    60 shuffles -> 165 ×16, 166 ×25, 167 ×19   canonical 167
dupmove,        30 shuffles -> 151 ×16, 152 ×14            canonical 152
oracle_subject, 30 shuffles -> 134 ×3, 135 ×12, 136 ×8, 137 ×7
SHIPPED oracle_name, 20 shuffles -> {158} every time
```

The shipped grader is order-invariant; iteration 3 is not. `oracle_name` — a guard the
document marks **MUST be 167** — reaches 167 in under a third of orderings.

The two order-sensitive emails are exactly the **G-5 nesting pairs**:

- `World_Cup_Cleat_Launch.reveal-event-date-and-venue-3` — ops `move 'Reveal Event'`
  (`{reveal}`) and `move 'Reveal Rehearsal'` (`{reveal, rehearsal}`). Both objects score
  1.0 for the first op; greedy takes op index 0 first and resolves the tie by pool
  position.
- `Partnership-with-deeptech-companies.caltech-conference-invitation` —
  `'FBS Planning Meeting'` (`{fbs}`) nested inside `'FBS Conference'` (`{conference, fbs}`).

A subset-scan over the corpus finds **5 distinct same-node same-kind nesting pairs**,
including one that is not merely nested but **identical**: `Company Retreat` and
`Retreat Company Meeting Call` both reduce to `{company, retreat}`. The document says the
`Company-Retreat` nesting *"resolves itself"* under turn membership. It resolves in
canonical order for this agent; the general case has no tie-break at all.

### 3.2 A perfect agent scores 165 through `sb.engine.run`

Not a synthetic shuffle — a permutation a real model would produce. Take the name-titled
perfect oracle (blocker 3's proposed `sb/oracle.py`), do moves as delete-then-recreate
(`sb/live/runner.py:109` explicitly permits it: *"delete-and-recreate is also fine"*), and
handle each email's ops in the reverse of answer-key order:

```
perfect agent, delete-then-recreate moves:
   ops in key order     : shipped 158   iter3 167
   ops in reverse order : shipped 158   iter3 165
Engine gate, name-titled oracle, node pool presented in reverse:
                          shipped 158   iter3 165
```

Two consequences:

- `sb/scale.py:126-127` prints `oracle: N/N = 100% (must be 100% — corpus is valid at
  scale)` and `CLAUDE.md` makes that a mandatory pre-flight. Under iteration 3 the gate's
  100% is contingent on the order `sb/oracle.py` happens to touch the store in, and on the
  order two emails appear in `corpus/nodes/World_Cup_Cleat_Launch.json`. A corpus edit that
  reorders those two obligations turns the gate red without changing anything semantic.
- `sb/grader.py:4-6` states the design contract: *"Grades one email against its answer key
  by inspecting the calendar/todo STATE (not the model's edit), so a reschedule done as
  update-in-place and one done as delete-then-recreate score identically
  (ANSWER_KEY_GRAMMAR.md §8)."* Iteration 3 breaks it. Two models with content-identical
  final state score 167 and 165.

VERIFY-phase2 claim 14 recorded determinism as **CONFIRMED** after 40 shuffles. That test
shuffled `real` and `shotgun7` only — the two worlds where it does not fire. It does not
transfer to iteration 3, or to iteration 2 (the `mine` tie-break is 0 for both objects
here, since both are stamped to earlier emails).

### 3.3 A cheap fix, measured

Assign obligations in order of **decreasing keyword-set size** (specificity first: give
`{reveal, rehearsal}` its object before `{reveal}` competes for it). Implemented as an
extra leading sort term in `ref3.grade_ref(specificity=True)`:

| | real | oracle_name | oracle_subject | oracle_inflect | null | dup5 | shot7/45/90 | wrongkind | dupmove |
|---|---|---|---|---|---|---|---|---|---|
| iteration 3 as designed | 114 | 167 | 137 | 159 | 64 | 64 | 64 | 69 | 152 |
| **+ specificity-first** | **114** | **167** | 135 | 158 | **64** | **64** | **64** | **69** | **152** |

Under 30 shuffles: `oracle_name` **167 ×30** and `dupmove` **152 ×30** (from 165–167 and
151–152). Cost: 2 on `oracle_subject`, 1 on `oracle_inflect`, 0 everywhere else. It does
not fix the `Company Retreat` identical-keyword pair — nothing in this contract can; that
one needs the G-5 name-aware lint and a corpus rename.

Greedy vs an optimal (max-passing-ops) assignment, for reference: greedy loses 2
create/move ops on `real` and 6 on `oracle_subject`. VERIFY-phase2 §3's argument against
switching to optimal still holds (optimal fabricates a pass on
`World_Cup_Cleat_Launch.tooling-po-needs-approval`); specificity-first recovers the
deterministic part of the gap without that risk.

---

## 4. Rule 4, the stale-survivor rule — the guard is tuned to one adversary

### 4.1 It is evaded by changing one word

`dupmove_retitle` (`adv.build(copies=0, retitle_stale=True)`): the model double-books
every move exactly as `dupmove` does, but drops the **last word** of the stale copy's
title, so the survivor scores below the `>= claimed score` floor.

```
dupmove plain      shipped 148   iter3 152     (guard threshold: <= 152)
dupmove retitled   shipped 146   iter3 163     FAILS the guard by 11
```

163 of 167 — four of fifteen double-bookings caught. The calendar still holds two entries
for every rescheduled obligation, on two different days. The `dupmove` guard therefore
certifies robustness against precisely one spelling of the adversary it was built from.
This is not a regression against the shipped grader (146 vs 148, both blind), but it does
mean the guard's "PASS" is worth much less than the document implies, and the guard as
written would not have caught it.

### 4.2 Its only firing on the real capture is a mis-diagnosis, and it flags two objects

`World_Cup_Cleat_Launch.reveal-event-date-and-venue-3`, `move 'Reveal Event'`, kws
`{reveal}` (`event` is a stop word). The document names one false extra. There are two:

```
claimed: event 'WC cleat reveal event'          2026-08-13   (reveal-event-date-and-venue)
extra  : event 'WC cleat press briefing'        2026-08-06   desc "…week before reveal"
extra  : event 'Design walk-through at WC reveal' 2026-08-13  desc "Design Lead speaking slot…"
```

Both extras are *other obligations' correct answers*. The second one matches **only
through its description** — the G-3 haystack the document keeps deliberately. The reason
string would read `moved, but 2 stale copies left behind (double-booked)`. With
`brake_hay='title'` the press briefing drops out and the reason becomes "1", which is why
the title-only variant scores identically but is not diagnostically identical.

The document's characterisation — *"a mis-diagnosis inside an email that fails anyway"* —
is **CONFIRMED**: the email's other op (`move 'Reveal Rehearsal'`) is on the wrong day
under both variants, and with `move_stale=False` the profile is `wrong day 30, stale 0`
with `failing ops` unchanged at 60. Stated plainly: **the stale rule's measured
true-positive rate on the real capture is 0 of 1**, and its whole justification is the
`dupmove` guard, which §4.1 shows is evadable. It should still ship — a benchmark that
cannot fail a double-booker is broken — but not on the evidence offered.

### 4.3 It costs a legitimately-perfect agent one email

`docs/grader-contract.md:167` records `iter. 3, no stale rule` at `oracle_subject` **138**
against iteration 3's **137**. That one point is a false stale flag on a
scheduling-perfect, subject-titling agent. The document does not mention it.

---

## 5. Rule 5, `cancel` through the assignment

### 5.1 The threshold is right, and stronger than the document claims

`nocancel` (`adv.build(skip_cancel=True)`): perfect in every way but never deletes
anything. iteration 3 = **159 = 167 − 8**, i.e. **all 8 cancel emails fail** and there are
**0 false-passing cancel ops**. τ = 1.0 is not loose on this corpus.

Better still, the historical **G-7** case the register calls the canonical failure —
`~"dynamics"` surviving as one conflated object in all four recorded runs — now passes:

```
kws {boston, dynamic, visit}
  'Boston tech trip (WHOOP + Boston Dynamics)'  overlap 0.67  -> cancel PASS
  'Boston trip - Boston Dynamics visit'         overlap 1.00  -> cancel FAIL
```

That is the single strongest piece of evidence for rule 5 and the document does not make
the argument.

### 5.2 Both remaining cancel failures — CONFIRMED, one more strongly than stated

`World_Cup_Cleat_Launch.manufacturing-kickoff-3` — `cancel 'Design Lead Stage Slot'`
(`match=['Design Lead Stage Slot']`, a G-8 whole-name default). The surviving object is
`"Design walk-through at WC reveal"`, description `"Design Lead speaking slot on stage at
reveal"`, **stamped `World_Cup_Cleat_Launch.reveal-event-date-and-venue-2`** — and that
email's answer key is
`[create 'Reveal Rehearsal', create 'Design Lead Stage Slot']`. So the surviving object is
the model's **own answer to the very obligation being cancelled**. The failure is not
merely genuine, it is unambiguous. The document's claim that the shipped grader **passed**
it also reproduces: `shipped passed=True ['cancelled']`, because no title contains the
literal phrase.

`Partnership-with-deeptech-companies.boston-dynamics-cancel` — the node pool at grading
time contains six separate events including `Boston trip - WHOOP HQ visit` (Aug 4) and
`Boston trip - Boston Dynamics visit` (Aug 6). The trip **is** split, so deleting the
Dynamics visit does not destroy the trip. **CONFIRMED**; a missed deletion.

### 5.3 What the corpus cannot test, and what breaks when it can

All 9 cancel ops live in 8 emails, and **all 8 are all-cancel emails**. There is not one
email in the corpus that mixes a `cancel` with a `create`/`move`. So rule 5's stated
benefit — *"A sibling op in the same email keeps the object it was supposed to keep"* — is
**UNFALSIFIABLE on this evidence**, and so is the failure mode it implies. Constructed
(`cases.py` case A), on a shape the schema permits:

```
ops: [create 'Alpha Beta' @serve, cancel 'Alpha Beta Gamma']
node already holds: event "Alpha Beta Gamma" on the serve date
the model does NOTHING
   iteration 3 : PASS  ['ok', 'cancelled']
   shipped     : FAIL  ['matched', 'should be cancelled, but 1 still on the calendar']
```

Phase 1 claims at floor 0.0 and unconditionally, so the create op claims the object that
was supposed to be deleted, which removes it from the cancel's view. A full pass for an
email where the model took no action at all. Latent today; it becomes live the moment a
phase-5 corpus edit puts a cancel and a create in one email.

### 5.4 The cancel is now cross-kind, which makes it strictly harder than the shipped rule

`cancel` claims *"the best unclaimed object, of either kind"*, and there is no kind check
on a cancel's claim (`phase2_guards.py:390-394`). The shipped grader checks only the
same-kind pool (`sb/grader.py:151-152`). Constructed (`cases.py` case C):

```
op: cancel 'Vendor Onsite' (event). Model deleted the event, left a to-do "Vendor Onsite prep".
   cross-kind pool : FAIL   kind-filtered : PASS
```

Costs nothing on this capture (kind-filtered scores identically on all 8 worlds), but it
is a hardening in the exact direction **G-7** says cancels are already too hard, and the
document never mentions it. K-7's four mis-keyed kinds are live exposure for it.

---

## 6. The cross-kind pool and the kind-first tie-break

### 6.1 "Score-neutral" — CONFIRMED as a measurement, OVERSTATED as a property

Verdict-level (not just score-level) comparison of the cross-kind pool against the
kind-filtered pool on `real`, `oracle_name`, `oracle_subject`, `oracle_inflect`,
`dupmove`, `wrongkind`, `dup5`, `shot45`: **0 email verdicts differ on any world**. So the
measurement is right, and stronger than the document states it.

It is not a property, though. There are two directions:

- **Cross-kind claiming can never turn a failing email into a passing one.** A wrong-kind
  claim always fails its op, and a cancel's claim always fails its cancel, so the email
  fails regardless. Proved by construction, not just measured. (`cases.py` case B shows
  the shielding happening — a spare event is claimed cross-kind by a to-do op and vanishes
  from the event op's duplicate check, flipping that op from `over-created` to `ok` — but
  the shielding op fails, so the email fails either way.)
- **It can turn a passing email into a failing one**, via the cancel's cross-kind reach
  (§5.4). That direction is real and unmeasured.

### 6.2 The stated reason for rejecting overlap-first is WRONG

> Overlap-first ordering was measured and rejected: it fails a model that creates both an
> event and a to-do for one obligation, which the shipped grader passes.
> — `docs/grader-contract.md`, rule 2

Measured (`adv.build(both_kinds=True)`, a perfect agent that creates an event **and** a
to-do for every obligation):

```
bothkinds   shipped 153   kind-first 166   overlap-first 166   kind-filtered 166
```

Both halves of the sentence fail. Structurally too: with both objects carrying the same
title they tie on overlap, and `kindok` is the next term in the overlap-first key, so the
right kind still wins. The scenario that *does* separate them is different — a
**badly-titled right-kind object plus a better-titled wrong-kind one**:

```
op create event 'Vendor Kickoff Call' (kws {vendor, kickoff});
model made event "Kickoff" (0.5) and todo "Vendor Kickoff" (1.0)
   kind-first PASS · kind-filtered PASS · shipped PASS · overlap-first FAIL 'wrong kind'
```

So the choice is defensible on a case the document does not describe — and note the
**kind-filtered pool gets the same protection**. Meanwhile the only guard column that
separates the orderings is `wrongkind`, where overlap-first sits at exactly the null floor
(64) and kind-first sits 5 above it (69). The document rejected the safer ordering citing
a measurement that does not reproduce.

### 6.3 The `wrongkind` guard's +5 — CONFIRMED mechanism, misattributed cause

Five emails where a *right-kind* sibling object shares a word and sits on the right day:
confirmed. But `kind-filtered` also scores **69**, so the +5 is not caused by the
cross-kind pool at all — it is the single-content-word identity weakness, and it is
present in every variant except overlap-first.

### 6.4 The 8 `wrong kind` claims — CONFIRMED as described, OVERSTATED in the changelog

All 8 reproduce, with the split the document gives. Six are genuine model-vs-key kind
disagreements (`Team_pizza_party`, `event day!`, `ask_about_patent_overlap`,
`retention_conversation`, `approve_trophy_correction`, `podcast_taping`). Two are
spurious claims of a *different obligation's* object at overlap 0.5:

- `World_Cup_Cleat_Launch.manufacturing-kickoff`, `create 'Manufacturing Kickoff'`
  (`{kickoff, manufactur}`) claims the to-do `"Sign off on WC cleat final colorway"`,
  whose description ends *"Tight against manufacturing."* The model created **nothing**
  for this obligation; the pool has no Manufacturing event. The true diagnosis is "not
  found"; iteration 3 confidently prints `wrong kind: created a to-do, expected a event`.
- `Sponsoring-Marathon.pitch-deck-2`, `create 'Pitch Breifing'` (`{breif, pitch}`) claims
  `"Ensure marketing creates marathon pitch deck"`. The obligation name contains a
  **typo** — no object will ever contain `breif`, so this op is structurally unpassable
  under rule 1 at better than 0.5. Reported as a kind error.

The 2026-08-29 changelog says the cross-kind pool *"turns 'nothing created' into `wrong
kind` when that is what happened (8 of 26 residual not-found ops)"*. It is what happened
in **6 of 8**; in 2 of 8 the new reason string is factually false. Since G-9 (better
failure attribution) is the point of the change, a 25% wrong-label rate should be recorded.

---

## 7. Identity: stop list, stemmer, and the keyword fallback

**Stop-word sensitivity — CONFIRMED to the digit.** The list has exactly **97** words;
**66** never occur in any op name; leave-one-out over all 97 against `real`,
`oracle_name`, `oracle_subject`, `oracle_inflect` moves some world for exactly **12**
words, never by more than ±2, and **none moves `oracle_name`**. The 12 are `call`,
`confirm`, `contact`, `create`, `day`, `final`, `go`, `list`, `meeting`, `of`, `review`,
`sync`.

**Defect: the stop list is applied *before* stemming, the haystack is stemmed.** So an
inflected form of a stopped word survives filtering and then stems back to the stopped
concept:

```
'Team Meetings'  -> {meet, team}      'Team Meeting' -> {team}
'Recap Notes'    -> {not, recap}
```

Four op names in the corpus are affected: `news_segment` (×2 emails — `news` → `new`,
and `new` **is** a stop word, so the obligation is half-satisfied by any object containing
"new"/"news"), `Send CTO Rubric from Meetings` (`meetings` → `meet`, matched by every
title containing "meeting"), and `add_promo_agenda_item`. Fix is one line —
`stem(w) not in STOPSTEM` — and it should be measured before shipping, not after.

**Defect: the stemmer over-strips `-es` and breaks the singular/plural pairing it exists
to create.** `notes → not` (colliding with the stop word `not`, which is *not* filtered
out of the haystack), `codes → cod`, `dates → dat`, `rates → rat`, `bases → bas`,
`gates → gat`, `sales → sal`, `series → ser`, `press → pres`, `atlas → atla`. `note`
stems to `note`, so `note`/`notes` do **not** unify — the exact inflection robustness rule
1 is justified by. 76 op-name words in this corpus change stem under `+s`; this is a large
part of `oracle_inflect`'s 8 losses. Guarding suffix stripping on a preceding consonant
(`-es` only after `s/x/z/ch/sh`) would remove most of it.

**Latent: an op name with no content and no alphanumerics.** `keywords_of` returns the
empty set, `overlap` returns 0.0, so a `create` can never be claimed (permanent "not
found") and a `cancel` can never be claimed (permanent **vacuous pass**, whatever is on
the calendar). `sb/schema.py:_parse_op` does not validate `name`. Not in this corpus;
worth a `CorpusError`.

**G-3, descriptions — the incentive has flipped sign.** Blanking every description in the
capture: shipped `97 → 93` (−4), iteration 3 `114 → 115` (**+1**). A perfect agent given a
realistic description (subject + first 180 chars of the body): shipped `158 → 142` (−16),
iteration 3 `167 → 164` (−3). So iteration 3 substantially reduces G-3's collision
penalty — real progress — but under it descriptions are now worth *minus* one point to the
real model, not plus four. Open issue 3's framing ("the cancel audit shows a case where
the description is what makes the grader right") is one-sided; §4.2 shows a case where a
description makes it wrong.

---

## 8. Blocker closures and the register's entry condition

| blocker | document's claim | verdict |
|---|---|---|
| kind filter unaddressed | cross-kind pool, kind-first, `wrong kind` reason; score-neutral | **PARTIAL** — score-neutral confirmed at verdict level on 8 worlds; not a property (§5.4, §6.1); the justification for kind-first is wrong (§6.2); 2 of 8 new labels are false (§6.4) |
| `cancel` bypasses the assignment | claims from what create/move leave unclaimed, at 1.0, either kind | **CONFIRMED** with two caveats: cross-kind makes cancel strictly harder (§5.4) and the sibling-shielding mechanism is a false-pass surface, untestable on this corpus (§5.3) |
| `sb/oracle.py:52` titles by `match` | becomes `op.name`, in the same commit; engine reads 167 / 163 / 158 | **CONFIRMED** — all three numbers reproduce; and 0 op names are substrings of a same-node same-kind sibling's name, so `Store.find_in_node` stays unambiguous under the new title policy |
| `sb/tests/test_e2e.py:58` would flip | it does not flip; reason `moved, but 1 stale copy left behind` | **CONFIRMED** — replayed through `sb.engine.run`: `beta.shift` fails with `stale copy on move`, `gamma.notice` fails over-acted, `alpha.brief`/`alpha.review`/`beta.book` pass. The 65-test suite passes unchanged on HEAD |

**Entry condition** (register changelog 2026-08-19: *"G-1, G-2, G-7 and the kind filter
are one contract"*) — **met in scope, not in substance.** All four are addressed in one
contract, which is the improvement over iteration 2. But G-2's replacement (the volume
brake) is defeated by §2, and G-7's canonical evidence (the conflated Boston trip) does not
exist in this capture, so the fix is verified only by construction (§5.1).

**The 148 → 158 correction — CONFIRMED.** The identity-tracking perfect agent scores 158
under the shipped grader; the iteration-2 double-booker scores 148. All 9 remaining losses
are the grader's, not the agent's: 2 are `Team_pizza_party` (the underscore name — G-1/G-8,
`match=['Team_pizza_party']` as one literal), 7 are `found N matching, expected exactly 1`
on distinct obligations (G-2). None is an agent fault.

**"Iteration 2's `oracle_engine` 166 was the dead tie-break, not a corpus ambiguity" —
OVERSTATED.** Iteration 3 does read 167, so the immediate claim holds. But "the
`Company-Retreat` name nesting resolves itself" is only true for one ordering: the nesting
is precisely what makes `oracle_name` unstable at 165–167 (§3.1), and `Company Retreat` /
`Retreat Company Meeting Call` have *identical* keyword sets, which no tie-break in this
contract can separate. The corpus ambiguity was real; it is masked, not resolved.

---

## 9. Downstream consumers — CONFIRMED, nothing breaks

Every consumer of `grade_email` in the repo, checked against `grade_v4`'s shape:

| consumer | reads | verdict |
|---|---|---|
| `sb/engine.py:147` | `grade_email(answer, ctx, state, turn)`, `.passed` | fine — signature unchanged |
| `sb/live/runner.py:612` | same, plus `_print_email` | fine |
| `sb/live/runner.py:82-87` (`_print_email`) | `d["passed"] / ["expected"] / ["actual"] / ["reason"]` | fine — `grade_v4` emits all four on every detail, including cancel and no-action |
| `sb/live/runner.py:625` (capture `verdicts`) | `asdict(EmailResult)` | fine — same dataclass |
| `sb/regrade.py:69` | same call, `.passed` | fine |
| `sb/analyze.py:25` | `re.compile(r"\b(PASS|FAIL)\b\s+\[(\d+)\]\s+(\S+)")` over the log | fine — never reads a reason string |
| `sb/tests/test_capture_regrade.py` | asserts live == offline under the *same* grader | fine — grader-agnostic. Note its `simulate()` appends on every op and never deletes on a move, so under iteration 3 it becomes a double-booker; the assertion still holds because both sides move together |
| `sb/tests/test_grader.py` | `details[0]["reason"] == "on the wrong day"` | fine — `grade_v4` emits exactly that string. (`keywords_of("filing")` = `{fil}`, matches) |
| `sb/schema.py` lint #5 | `op.match` | dead but harmless, as the document says |
| `webapp/` | — | no reference to `grade_email` or `EmailResult` outside the generated `webapp/api/_lib/sb/` |

`.venv/bin/python -m pytest sb/tests -q` → **65 passed** on HEAD (1f95cbc, clean tree).
`.venv/bin/python -m sb.scale --filler 0 --seed 42 --days 200` → `oracle: 167/167 = 100%`
on HEAD.

---

## 10. Verdict table

| # | claim | verdict |
|---|---|---|
| 1 | shipped 97 / 167 / 158 / 95 / 142 / 64 / 64 / 64×3 / 65 / 148 | **CONFIRMED** (independent harness; prototype agrees) |
| 2 | iteration 3 = 114 / 167 / 167 / 137 / 159 / 64 / 64 / 64×3 / 69 / 152 | **CONFIRMED** |
| 3 | all six variant rows (cancel any-word, ≥0.5, no-stale, kind-filtered, overlap-first, title-only) | **CONFIRMED**, every cell |
| 4 | engine-path 167 / 163 / 158 for the three oracle-title permutations | **CONFIRMED** |
| 5 | failure profile 56/50/—/14/11/—/3/2/80 → 76/18/8/29/0/1/2/2/60 | **CONFIRMED**, every cell |
| 6 | 0 email flips against iteration 2 | **CONFIRMED** |
| 7 | `dupmove` = 152 = 167 − 15, every double-booked move fails and nothing else does | **CONFIRMED** in canonical order; guard is linear and sensitive. **UNSTABLE** — 151 in 16 of 30 pool shuffles |
| 8 | shipped's `dupmove` 148 is a pass (✓, "fails `oracle_name`" singular) | **WRONG** — the prototype's own checker prints `FAIL -> oracle_name 158, dupmove 148`; 148 > 143 |
| 9 | the shipped grader cannot award full marks (158), the 9 losses are the substring rule's | **CONFIRMED** — 2 G-1/G-8, 7 G-2, 0 agent faults |
| 10 | turn membership "cannot be defeated by a mis-stamped `email_id` (A-5)" | **WRONG / UNSAFE** — laundered duplicates score **167** vs shipped 72; 509 objects vs 99 |
| 11 | "`email_id` is not consulted at all now" | **WRONG** — on `runner.py:611` / `regrade.py:68` it *defines* turn membership, which rules 3 and 4 depend on |
| 12 | iteration 3 removes the disagreement between the two grading paths | **WRONG** — the same state now scores 64 (engine) vs 167 (runner) |
| 13 | `oracle_engine` and `oracle_name` MUST be, and are, 167 | **UNSAFE** — 167 in 19 of 60 pool orderings; a perfect agent handling two moves in the other order reads 165 through `sb.engine.run` |
| 14 | grading is state-based (`sb/grader.py:5-6`) | **WRONG under iteration 3** — verdict depends on store list order, i.e. on the model's edit sequence |
| 15 | kind-first is score-identical to the kind-filtered pool | **CONFIRMED** as a measurement (0 verdict flips, 8 worlds); **OVERSTATED** as a property (§5.4 constructed counter-case) |
| 16 | overlap-first "fails a model that creates both kinds, which the shipped grader passes" | **WRONG** — 166 / 166 / 166 / **shipped 153**; the separating case is a differently-shaped one |
| 17 | the `wrongkind` +5 is right-kind siblings on the right day | **CONFIRMED** mechanism; **misattributed** — kind-filtered is also 69, so the cross-kind pool is not the cause |
| 18 | 8 wrong-kind claims, 6 genuine + 2 weak | **CONFIRMED**; changelog wording ("when that is what happened, 8 of 26") **OVERSTATED** — true for 6 of 8 |
| 19 | both remaining cancel failures are genuine; one was a false pass under shipped | **CONFIRMED**, both, with the object and stamp identified; the `manufacturing-kickoff-3` case is stronger than stated |
| 20 | cancel τ = 1.0 is tight | **CONFIRMED** — `nocancel` fails all 8 cancel emails, 0 false-passing cancel ops. Also fixes the historical G-7 Boston case (0.67 < 1.0) |
| 21 | rule 5 lets a sibling op keep the object it was supposed to keep | **UNFALSIFIABLE** on this corpus (all 8 cancel emails are all-cancel) — and the constructed case is a **false PASS** where shipped fails |
| 22 | cancel is unchanged in the kind dimension | **not claimed, and UNSAFE** — cancel is now cross-kind, strictly harder than `sb/grader.py:151` |
| 23 | the one stale flag is a mis-diagnosis in an email that fails anyway | **CONFIRMED** — and the stale rule's true-positive rate on `real` is 0/1; it flags **2** objects, not 1, one of them via its description |
| 24 | the stale rule is what stops `dupmove` reaching 167 | **CONFIRMED**, but **evadable** — dropping one word from the stale copy's title scores 163 |
| 25 | the stale rule is free on the oracles | **WRONG** — it costs `oracle_subject` 1 (138 → 137); the document's own table shows it and the text does not |
| 26 | title-only brake changes no verdict on any world | **CONFIRMED** at verdict level (1 op-reason difference on `real`) |
| 27 | stop-word list: 97 words, 66 unused, 12 move a world by ≤ ±2, none moves `oracle_name` | **CONFIRMED** to the digit |
| 28 | the stop list is sound as written | **OVERSTATED** — it is applied pre-stem while the haystack is stemmed, so `news`→`new`, `meetings`→`meet` leak on 4 op names; and the stemmer breaks `note`/`notes`, `date`/`dates`, `code`/`codes` |
| 29 | `test_e2e.py:58` does not flip, reason `stale copy on move` | **CONFIRMED** by replay |
| 30 | `sb/oracle.py` can safely title by `op.name` | **CONFIRMED** — 0 same-node same-kind name-substring collisions, so `find_in_node` stays unambiguous |
| 31 | downstream consumers are unaffected | **CONFIRMED** — all 10 checked |
| 32 | the register's phase-2 entry condition is met | **PARTIAL** — all four addressed in one contract, but G-2's replacement is defeated (§2) and G-7's canonical case is absent from this capture |
| 33 | descriptions in the haystack are kept deliberately and are net-positive | **OVERSTATED** — under iteration 3 blanking every description *raises* `real` from 114 to 115; the G-3 collision penalty on a perfect agent does drop from −16 to −3, which is genuine progress |
| 34 | greedy assignment is adequate | **CONFIRMED** with a caveat — greedy loses 2 create/move ops on `real`, 6 on `oracle_subject`; optimal still fabricates a pass (VERIFY-phase2 §3), so do not switch |
| 35 | the guard set is built from the adversaries that broke iteration 1 | **CONFIRMED**, and still insufficient — every world stamps `email_id` correctly and presents objects in creation order, which is why §2 and §3 are invisible to it |

---

## 11. What I could not check

1. **Whether any of this is closer to the truth.** Unchanged from VERIFY-phase2 §7.1.
   Phase 1d is deferred behind C-10, so every number here is grader-versus-grader. My
   audits in §4.2, §5.2 and §6.4 are my own judgement of the capture, not a human
   reference.
2. **Generalisation.** One capture, one model, one seed (42), one lever set (1/5/7), one
   corpus sha `03e0d963b9866d8f`. K-2 records 19 of 100 seeds raising `InfeasibleSchedule`,
   so seed variance cannot be run. The +17 on `real` could be noise of the same size as the
   91→97 churn the register attributes to nondeterminism.
3. **Whether a real model would launder an `email_id`.** A-5 says there is no observed
   instance in any run, and A-3's counter-evidence stands: where models over-acted they
   stamped correctly. §2 is an exploit that exists, not one that has been used. Note the
   direction of the risk though: it is not a model that games it deliberately, it is a
   model that copies the wrong id from a long list (A-4: 108 of 167 ids exceed 40 chars,
   16 prefix pairs) and is silently rewarded.
4. **Predicate coverage.** Unchanged: 112 `eq`, 12 `by`, 1 `any_of`, 0 `in`/`not_in`, and
   `exact_day` on 134 of 134 ops. Nothing here tests interval predicates or `within:Nd`.
5. **Live-store ordering.** §3's shuffle is a proxy. I did not read `store_app.py`'s
   `/state` ordering guarantees; if it is anything but stable insertion order, §3 becomes a
   live nondeterminism rather than a fragility. That is worth ten minutes before
   implementation.
6. **The corpus authority question.** C-10 is resolved *for working purposes*; the contract
   keys identity on `op.name`, so if the outstanding yes/no with the corpus authors goes
   the other way every number here moves.
7. **`sb/live/mcp_app.py` / `store_app.py` behaviour on duplicate creates.** If the store
   already rejects or warns on an exact-duplicate create, §2's adversary is less reachable
   than it looks. `_watch_attribution` (A-4) explicitly does not warn on a same-day sibling
   id, which is what §2 uses.

---

## 12. Recommendations for the implementation, if it proceeds

Not a verifier's decision, but these follow directly from the numbers.

1. **Close the laundering hole before anything ships. I measured the fix and it is free.**
   Scope the volume brake (and the stale rule's freshness test) to the **day's** new-object
   set rather than the email_id-filtered turn: `sb/live/runner.py:605` already computes
   `day_new` and `sb/regrade.py:63` already reads it, so it is one extra argument at
   `runner.py:611` / `regrade.py:68`, and it is a no-op on the engine path (where the
   per-email delta already *is* the turn, `sb/engine.py:145`). Measured with `ref3.py`'s
   `day_brake` mode:

   | | real | oracle_name | oracle_subject | oracle_inflect | null | dup5 | shot7/45/90 | wrongkind | dupmove | `launder` | `launder_all` |
   |---|---|---|---|---|---|---|---|---|---|---|---|
   | iteration 3 as designed | 114 | 167 | 137 | 159 | 64 | 64 | 64 | 69 | 152 | **167** | **152** |
   | **brake scoped to the day** | **114** | **167** | **137** | **159** | **64** | **64** | **64** | **69** | **152** | **77** | **64** |

   Identical on every guard column and on the real capture; `launder` falls from 167 to 77
   (below the shipped grader's own oracle and 13 above the floor — those 13 are the move
   emails, where the agent genuinely leaves no duplicate) and `launder_all` falls to exactly
   the null floor. The theoretical cost — a same-day sibling's correct object being read as
   a duplicate, i.e. G-2 returning — does not fire once on any world measured here, but it
   is the thing to watch. Whatever is chosen, **the fix must be guarded**: add a `launder`
   world to `phase2_guards.WORLDS` that stamps duplicates with a node sibling's id, and
   require it not to exceed `null` + the move-email count. Without a guard this hole is
   reachable again in one refactor.
2. **Add a deterministic tie-break, and prove it with a shuffle test.** Specificity-first
   (assign obligations in decreasing keyword-set size) makes `oracle_name` 167 and
   `dupmove` 152 in 30 of 30 shuffles at a cost of 2 on `oracle_subject` and 1 on
   `oracle_inflect`. Whatever is chosen, `sb/tests/test_grader_guards.py` should shuffle the
   pool order N times and assert the score is invariant — otherwise the mandatory
   `sb.scale` gate is a coin flip and `sb/grader.py:5-6`'s stated contract is false.
3. **Keep the kind filter on the create/move assignment; add `wrong kind` as a
   non-claiming second pass.** It scores identically on every measured world, it removes
   the cross-kind cancel over-reach (§5.4) and the duplicate shielding (§6.1), and it makes
   the overlap-first-vs-kind-first argument moot. If cross-kind claiming is kept anyway,
   fix the rationale in the doc — the one given does not reproduce (§6.2) — and **exclude
   wrong-kind objects from the cancel phase**.
4. **Do not present the `wrong kind` reason when the claim is weak.** Require the
   cross-kind claim to reach some floor (0.75 would leave all 6 genuine cases and drop both
   false ones) or say `no <kind> titled like "X" (closest: a <other-kind> …)`. As written,
   2 of 8 new labels are wrong, which is a step backwards for G-9.
5. **Fix the stop-list/stem ordering** (`stem(w) not in {stem(s) for s in STOP}`) and
   re-run the leave-one-out. Guard the `-es` strip on a preceding sibilant. Both are
   one-liners with measurable effects and both should be measured before, not after.
6. **Raise on an empty keyword set** in `sb/schema.py` — a cancel of such an obligation
   passes vacuously whatever is on the calendar.
7. **Rename the `dupmove` guard's requirement, or add the retitled variant.** As written
   it certifies robustness against one spelling; the retitled double-booker scores 163
   against a 152 threshold.
8. **Correct three things in `docs/grader-contract.md`** regardless of what ships: the
   shipped row's `dupmove` ✓ (line 161), the overlap-first rationale (rule 2), and the
   claim that turn membership is immune to `email_id` (rule 2, and both changelog entries).
9. **Per the register's status legend, this stays `fix proposed`.** With items 1 and 2
   done, and re-verified, it is implementable. Without them it should not be written into
   `sb/grader.py`, because both defects would then be baked into the number every future
   run is compared against.
