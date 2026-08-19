# VERIFY — red-team of the phase-2 identity contract (`grade_email_v2`)

Adversarial pass over the **proposed, not-yet-implemented** replacement for the grader's
identity contract. **This file reports; it changes nothing.** No edit to `sb/`, `corpus/`,
`captures/`, the prototype, or any other doc. No live model run. `scripts/recover_corpus.py`
and `scripts/fix_match.py` were neither read nor executed.

Prototype under test:
`/private/tmp/claude-501/-Users-jamesoc-dev-SecretaryBench/fe67c85a-9c16-441a-99a9-a4b047c35b38/scratchpad/phase2.py`
(`grade_email_v2`, `keywords_of`, `score_v2`, `eval_v2`).

**Method.** I wrote my own scoring harness from the spec rather than calling `eval_v2`, so a
bug in the prototype's measurement code could not hide. It re-implements keyword extraction,
word-level matching, exclusive assignment, the attribution tie-break, node/kind pool
construction, the date predicate and the day loop. It shares nothing with `phase2.py` and
borrows from `sb/` only `resolver.resolve` (date arithmetic, identical under both contracts).
Scratch code lives outside the repo at
`…/scratchpad/ref.py`; every command below is prefixed
`PYTHONPATH=/Users/jamesoc/dev/SecretaryBench:…/scratchpad .venv/bin/python`.

---

## Bottom line

**42 discrete claims checked: 27 CONFIRMED, 2 WRONG, 6 OVERSTATED, 7 UNSAFE.**

Every headline number reproduces **to the digit** under an independent implementation, and
the prototype's own `eval_v2` agrees with mine exactly. The measurement is sound. **The
safety argument is not.** Two of the three guards do not hold, and the reason they appear to
hold is that the guard set contains no adversary that varies the one thing the new contract
stopped constraining: **how many objects the model creates.**

| | shipped | proposed |
|---|---|---|
| real (1c capture) | 97/167 | **112/167** |
| the proposal's oracle (title = `op.name`, duplicate on move) | 148/167 | **167/167** |
| **`sb/oracle.py` — the repo's actual reference agent** | **167/167** | **161/167** |
| null | 64/167 | 64/167 |
| adversarial +3 days | 64/167 | 64/167 |
| **adversarial: right answer, plus 5 duplicates of every object** | **64/167** | **167/167** |
| **adversarial: never reads a date, creates one object per day for 7 days** | **56/167** | **93/167** |
| **adversarial: same, 45-day spread** | **56/167** | **128/167** |
| **adversarial: same, 90-day spread** | **56/167** | **148/167** |

The last three rows are strategies that **beat the null floor under the proposed contract and
fall below it under the shipped one**. That is the regression criterion stated in the brief,
met three times over.

**The single most important problem:** the contract has **no volume brake**. Deleting the
`count_ok` rule removed the only penalty for creating extra objects, and the `dateok` term in
the assignment tie-break then turns grading into an active *search* for an object that
satisfies the date. Together they mean a model that knows the vocabulary but not the calendar
can buy the score with volume. A date-blind shotgun scores **148/167 (89%)** — thirty-six
points above the real model this repair is measuring, and ninety-two above the same agent
under today's grader.

The good news, and it is substantial: **the +15 on the real capture is genuinely earned** (I
audited all 16 flips object-by-object against the capture; zero are for a bad reason), and the
`dateok` half of the hole can be removed at **zero cost to any measured world**.

---

## 1. Reproducing the four numbers independently — CONFIRMED, exactly

```
PYTHONPATH=… .venv/bin/python -c "import ref; …"      # ref.run(DAYS,'shipped'|'proposed')
```

| contract | real | oracle | null | adversarial |
|---|---|---|---|---|
| shipped, claimed | 97 | 148 | 64 | 64 |
| **shipped, mine** | **97** | **148** | **64** | **64** |
| proposed, claimed | 112 | 167 | 64 | 64 |
| **proposed, mine** | **112** | **167** | **64** | **64** |

Op-level: shipped 110/190, proposed 128/190. Baseline sanity:
`.venv/bin/python -m sb.regrade captures/baseline-sonnet-4-5` → `RESCORE 97/167`,
`as-run 97/167 (identical)`.

**The claimed failure-profile shift — CONFIRMED, every cell.** My op-level transition matrix
(shipped bucket → proposed bucket, n=190):

```
ok               -> ok                55        not_found  -> not_found  26
no-action:pass   -> no-action:pass    54        not_found  -> wrong_day  13
wrong_day        -> wrong_day         13        not_found  -> ok         11
count_too_many   -> ok                 7        count_too_many -> wrong_day  4
cancel_residue   -> cancel_residue     3        wrong_day  -> ok           1
no-action:over-acted -> same           2        ok         -> cancel_residue 1
```

→ `not found` 50→26, `wrong day` 14→30, `count: too many` 11→0, cancel residue 3→4,
over-acted 2→2, failing ops 80→62. All as claimed.

**The prototype's own measurement code is clean.** `phase2.eval_v2('proposed')` returns
`{real 112, oracle 167, null 64, adversarial 64}` and `eval_v2(strict_all=True)` returns
`{95, 167, 64, 64}`; my independent harness returns 112 and 95 for the same two. No
measurement bug found in `phase2.py`.

**One arithmetic caveat, OVERSTATED.** "The +15 emails" is a *net*. There are **16** emails
that flip fail→pass and **1** that flips pass→fail
(`World_Cup_Cleat_Launch.manufacturing-kickoff-3`, §5). Describing the change as "+15 emails"
hides a regression that a reviewer should see.

---

## 2. The guards

### 2.1 The oracle guard — **UNSAFE, and circular**

The oracle titles every object `op.name.replace('_',' ')`, which is *exactly* the string
`keywords_of` derives identity from. It is the one title policy in the space that cannot fail.
Sweeping the policy, scheduling held perfect in every variant:

| perfect agent, title policy | shipped | proposed |
|---|---|---|
| `" ".join(op.match)` — what `sb/oracle.py` actually writes | 153 | 162 |
| **`op.name` humanised — the proposal's oracle** | 148 | **167** |
| the email's subject line | 95 | 135 |
| subject + humanised name | 145 | 165 |
| name minus its last word | 111 | **167** |
| name, hyphen-joined | 148 | **167** |
| **name, every word inflected (`trophy`→`trophies`)** | **127** | **68** |

Three things fall out.

**(a) 167 is a property of the oracle, not of the contract.** Change the title policy at all —
even to the answer key's own `match` keywords — and it is gone. There is **no** title policy
that scores 167 under both contracts.

**(b) The repo's real oracle *loses* six emails.** Run `sb/oracle.py` through the real
`sb.engine.run` with the v2 contract monkeypatched over `sb.engine.grade_email`:

```
--days  57: sb/oracle.py through sb.engine.run  ->  shipped 167/167   PROPOSED 161/167
--days 200: sb/oracle.py through sb.engine.run  ->  shipped 167/167   PROPOSED 161/167
```

The six: `Company-Retreat.tasks-before-the-retreat`, `Company-Retreat.inquiry-on-vip-list`
(oracle titles the object `list`, which is a **stop word**, so `keywords_of` = `['vip']` and
nothing matches); `Partnership-with-deeptech-companies.spadxtech-meeting-before-fbs`
(`match=['spad']`, `keywords_of`=`['spadxtech']` — word-level matching cannot see a
sub-token); and three `cancel` ops (§4.2).

**This is implementation-blocking, and the proposal does not mention it.**
`sb/scale.py:126-127` prints `oracle: N/N = 100% (must be 100% — corpus is valid at scale)`,
and `CLAUDE.md` makes `.venv/bin/python -m sb.scale --filler 0 --seed 42 --days 200
--dst build/scaled0` a **mandatory pre-flight that must print 100%**. Shipping `grade_email_v2`
without also changing `sb/oracle.py:52` turns that gate red and makes the corpus look invalid.
The grader and the oracle are a coupled change.

**(c) Word-level matching is *less* robust than the substring rule it replaces.** A perfect
agent that pluralises — `trophy`→`trophies`, `advertising`→`advertisements` — scores
**68/167 under the proposal against 127/167 under the shipped grader**, four points above the
do-nothing floor. Substring matching absorbs morphological variation for free (`trophy` ⊂
`trophys`); a word-set does not. The proposal's stated motivation is robustness to how a model
words a title; on this axis it is a **regression**, and it is unguarded because the oracle
never inflects anything. A harder oracle (subject-titled) scores 135 — the honest ceiling for
a model that does not know obligation names.

### 2.2 The null guard — **CONFIRMED, but vacuous**

64/167 under both contracts, and 64 = 56 emails with no ops + 8 emails whose ops are all
`cancel` (a cancel passes when the pool is empty). This guard is **structurally invariant**:
any contract that (i) fails an op with zero objects and (ii) passes a no-action email with
zero creations scores exactly 64. It cannot discriminate between candidate contracts and
should not be counted as evidence for one. Every variant I tested — including the ones that
score 148 for a date-blind adversary — scores exactly 64 here.

### 2.3 The adversarial guard — **UNSAFE**

`+3 days` is the weakest adversary in its class. The corpus has **112 `eq`, 12 `by`, 1
`any_of`, 0 `in` predicates and `exact_day` tolerance on 134/134 ops**, so *any* uniform date
shift ≥1 day collapses to the floor under *both* contracts and can never separate them:

```
all dates +1d  shipped 65  proposed 65      all dates +5d   shipped 64  proposed 64
all dates +2d  shipped 65  proposed 65      all dates +10d  shipped 64  proposed 64
everything on the serve date  shipped 63  proposed 65
```

The adversaries that *do* separate them vary **count**, which the guard set never does:

| adversary | shipped | proposed | Δ vs null |
|---|---|---|---|
| right title, right day, **5 duplicates of every object** | 64 | **167** | **+103** |
| right title, **wrong kind** | 65 | 69 | +5 |
| right everything, **`email_id` of a sibling email** | 148 | 166 | +102 |
| **date never read**, one object/day for 7 days from serve | 56 | **93** | **+29** |
| same, 30-day spread | 56 | **116** | +52 |
| same, 45-day spread | 56 | **128** | +64 |
| same, 90-day spread | 56 | **148** | **+84** |
| same, but one title stuffed with **every** op name of the email | 56 | 127 | +63 |
| same, titles taken from the **email subject** (no answer-key text) | 59 | 109 | +45 |

**The duplicate hole.** `sb/grader.py:11-13` states the contract the proposal deletes —
*"exactly ONE object matching the obligation … a reschedule that leaves a stale duplicate
fails"* — and `runner.py:107` tells the model *"Never leave duplicates"*. Under the proposal
that instruction is unenforceable: an agent that does the right thing **and** leaves four
stale copies of every object scores a perfect 167. There is a named artifact for this:
`sb/tests/test_e2e.py:58` asserts `not res.results["beta.shift"].passed  # double-booked`.
Replaying that test with the v2 contract patched over `sb.engine.grade_email`:

```
SHIPPED   beta.shift (double-booked reschedule) passed = False   score 0.50
PROPOSED  beta.shift (double-booked reschedule) passed = True    score 0.67
```

**An existing test in the shipped suite encodes the rule the proposal reverses, and it fails.**
The 65-test suite passes on `main` (`.venv/bin/python -m pytest sb/tests -q` → `65 passed`).

**The shotgun hole, and its removable cause.** `grade_email_v2` scores each (op, object) pair
by `(overlap, mine, dateok)` and sorts descending. The third term makes the assignment prefer,
among equally-titled objects, *the one that satisfies the date* — i.e. the grader searches the
model's output for a date-satisfying candidate on the model's behalf. Removing it costs
nothing anywhere and closes most of the hole:

| variant | real | oracle(name) | `sb/oracle.py` | null | +3d | dup×5 | shotgun 7d | shotgun 45d |
|---|---|---|---|---|---|---|---|---|
| shipped | 97 | 148 | 167 | 64 | 64 | 64 | 56 | 56 |
| **proposed as designed** | 112 | 167 | 162 | 64 | 64 | **167** | **93** | **128** |
| **proposed, no `dateok` tie-break** | **112** | **167** | **162** | **64** | **64** | 167 | **65** | **65** |
| proposed, `strict_all` words | 95 | 167 | 88 | 64 | 64 | 167 | 93 | 128 |

**Dropping `dateok` is free on every world that was measured and removes 63 points of
adversarial headroom.** The duplicate hole is orthogonal and needs a separate brake.

---

## 3. The assignment

**Greedy vs optimal — the objection does not survive contact with the evidence.** I
implemented a maximum-cardinality bipartite matching that maximises the number of ops assigned
a *date-satisfying* object, and compared per email:

| world | greedy | optimal | Δ |
|---|---|---|---|
| real (1c capture) | 112 | 113 | +1 |
| oracle (name) | 167 | 167 | 0 |
| `sb/oracle.py` | 162 | 162 | 0 |
| oracle (subject) | 134 | 134 | 0 |
| shotgun 7d | 93 | 96 | +3 |

Greedy is provably suboptimal in general (I can construct the case: two ops whose keyword sets
nest, competing for one object, resolved by op index). But the single divergence on real data
— `World_Cup_Cleat_Launch.tooling-po-needs-approval` — goes the *other* way. The op is
`create 'Approve tooling PO'` (`kw=['approve','tooling','po']`, `eq this:FRI` → 2026-07-17).
The model created `'Approve tooling PO for factory'` due **2026-07-19**, which is genuinely the
wrong day. Optimal assignment passes the email by handing the obligation
`'Approve revised WC reveal event budget'` (due 2026-07-17, created three days earlier for a
*different* email, overlap 1/3 on the word "approve"). **Greedy avoids a false positive that
optimal assignment creates.** Verdict on my own attack: **OVERSTATED** — do not switch to
Hungarian/optimal; the higher score it buys is fabricated.

**Determinism — the objection does not hold.** `pairs.sort` is stable and the sort key is a
total order on `(overlap, mine, dateok)`; residual ties resolve by **list** insertion order
(op index, then pool index), never by dict iteration. Shuffling the pool order 40 times on the
real capture and 10 times on the shotgun world:

```
real, proposed, pool order shuffled x40 -> [112]   canonical 112
shotgun7,       shuffled x10            -> [93]    canonical 93
```

No flips. Not Python-version-sensitive. **CONFIRMED deterministic.** Two nits: `claimed_obj`
is keyed on `id(o)`, which is safe only because every pool object is alive for the duration of
the call and `_node_state` builds a fresh `Obj` per record — fine today, brittle if the pool is
ever built by reference; and the tie-break is *total* but *arbitrary* at the last step (op
order = answer-key order), which is an authoring artifact, not a semantic one.

---

## 4. The identity rule

### 4.1 Single-word obligations — real hazard, **does not fire once** on this evidence

**24 of 134 ops (18%) reduce to a single content word**, several of them generic: `board`,
`budget`, `people`, `pitch`, `ai`, `sponsor`, `recap`, `reveal`, `vision`. Candidate slack
roughly doubles: mean candidate objects per create/move op **0.72 → 1.30**, ops with zero
candidates **50 → 25**, max 5.

I tested exploitability directly by ablation — for each of the **69** create/move ops passing
under the proposal on the real capture, delete the object it claimed and re-grade:

```
of 69 passing create/move ops, still pass after deleting the object they claimed: 0 (0%)
```

**Every pass is uniquely earned.** The date predicate is doing the work: an unrelated object
must also land on the exact right day. So the "one common word lets anything through" attack
is real in principle and **inert in practice — on a model that creates 81 objects**. It is
inert only because the volume is low, which is precisely what the shotgun adversary removes.
The identity rule has no defence in depth; the date predicate is its only defence, and §2.3
shows how to overwhelm it.

**The all-stop-word fallback (`event day!`) — latent, unfired.** Exactly one op in the corpus
takes it (`Sponsoring-Marathon.approval-of-event`), and `keywords_of` returns `['event','day']`
— two words that are stop words for *every other* op, so this obligation alone is satisfied by
any object whose title or description contains "event" or "day" on the right date. It fails
under both contracts today only because the todo pool is empty when it is graded. It is a live
gaming surface if the corpus ever grows another such name.

**Keyword stuffing — caught by the assignment, defeated by volume.** One object per email
titled with every obligation name concatenated scores **127/167** when combined with a date
shotgun (§2.3), but exclusivity does its job when volume is normal: an email with two ops and
one stuffed object can satisfy only one of them.

**A defect with no measured effect.** `keywords_of` returns a **list**, so a repeated token is
double-weighted. `manufacturing-kickoff-2`'s `'Design Lead 1:1'` yields `['design','lead','1','1']`;
a title containing only "1" scores 2/4 = 0.50 while one containing "design" scores 1/4 = 0.25,
so the junk token outranks the real one. De-duplicating leaves the real capture at 112, so
this is latent, not live. Same class: the possessive splits into a one-character token
(`"Double check Josh's advertising work"` → `['double','josh','s','advertising','work']`).

### 4.2 `cancel` is widened, and G-7 gets worse — **UNSAFE**

`grade_email_v2` grades `cancel` by **any-word** absence over the cumulative node pool. That is
strictly wider than the shipped rule (all `match` keywords as substrings), so cancels get
*harder*, against a register finding (**G-7**) that says they are already too hard. Measured
consequences: cancel residue on the real capture 3 → 4; on `sb/oracle.py`, **3 of its 6 new
failures are cancels**; the one pass→fail email is a cancel (§5).

### 4.3 Scope items

- **`move` semantics — CONFIRMED sound.** All three `move` ops in the flip set claim an object
  inherited from an earlier email in the same node, which is exactly the intent. The attribution
  tie-break behaves as designed.
- **Attribution is a tie-break, not a filter — CONFIRMED, and it protects nothing.** An agent
  that is right in every way but stamps a sibling email's `email_id` scores **166/167** (shipped
  148). A-5's exposure is unchanged by this contract; do not record it as addressed.
- **No-action emails — unchanged.** over-acted 2 → 2. The no-action rule is now the *only*
  remaining brake on over-action, and A-3 lists five ways past it.
- **Objects created after their email's serve day — not an issue in this evidence.** 0 of 81
  captured objects were first seen on a day later than their email's serve day, so one-shot
  grading costs nothing here. Unchanged by the proposal.
- **Emails with 3+ ops:** 4 emails have 3 ops, 1 has 4. shipped 0/4 → proposed 1/4 on the
  3-op emails. 1-op emails 41→53, 2-op 2→4.
- **G-6 (cumulative pool decay) is attenuated, not fixed.** Pass rate by prior same-kind
  obligations in the node: shipped `0: 54% · 1-2: 45% · 3-5: 26% · 6+: 33%`; proposed
  `0: 62% · 1-2: 64% · 3-5: 44% · 6+: 48%`. The gradient survives.

---

## 5. Is the +15 real? — **CONFIRMED, all 16 flips**

I extracted, for every flipped email, the object the greedy assignment actually claimed (my
first attempt sorted by overlap alone and mis-identified two claims; the corrected extraction
replays the full `(overlap, mine, dateok)` key). Result:

- **16 emails flip fail→pass. 0 of them do so for a bad reason.**
- **create-ops in the flip set satisfied by an object attributed to a different email: 0.**
  Every create claims an object the model created *for that very email*.
- The three inherited claims are all `move` ops (`whoop-meeting-reschedule`, `demo-moved`,
  `inquiry-on-vip-list`) — correct semantics.
- Only **1 of 20** flipped ops claims at overlap < 0.5:
  `"Double check Josh's advertising work"` → `"Review Josh's marketing advertisements"` (0.40,
  matched on `['josh','s']`), and inspection of the capture shows that is plainly the right work
  on the right day.
- Spot-checks against `state_final.json` confirm the two flips I initially suspected were
  fabricated are in fact exact: `pitch_final` ← `"Innovation pitch comp final"` **2026-07-15**
  (answer key `@pitch_final` = 2026-07-15); `LeBron James marketing campaign scheduled` ←
  `"LeBron James - marketing campaign"` **2026-08-24** (key = 2026-08-24).

**The one regression.** `World_Cup_Cleat_Launch.manufacturing-kickoff-3`: a lone
`cancel 'Design Lead Stage Slot'`. The model correctly created nothing. Shipped passes it
(no object contains the literal phrase). Proposed **fails** it, because `['design','lead',
'stage','slot']` any-word-matches `'Design walk-through at WC reveal'` — an object created for
a *different* email that the model was supposed to keep. **A false negative introduced by the
new cancel rule**, and the exact shape of G-7.

**`count: too many` 11→0 — CONFIRMED as real elimination, with a caveat.** The bucket vanishes
because the rule is deleted, so "0" is true by construction rather than by diagnosis. Of the 11
ops: **7 → `ok`, 4 → `wrong_day`**. None is silently reclassified into a *different* false
failure. The 7 that pass are consistent with G-2's finding that 0 of 57 historical "duplicate"
failures were real duplicates — but note the proposal generalises that measurement into a
permanent rule, and the dup×5 adversary (§2.3) shows the generalisation is strictly stronger
than the evidence supports.

---

## 6. Effects on findings the register records as measured

- **G-2's "0 of 57 duplicates are real" — meaning preserved, scope widened.** The measurement
  is about four historical logs and is untouched. But the register's *option* list for G-2 offers
  "best-match assignment" and "relax to ≥1 satisfying, **no stale survivor**". The proposal
  implements the first and **drops the second**, which is the half that kept duplicate detection
  alive. That is a scope change the design does not flag.
- **§4.3's 76 / 14 / 10 / 0 split — invalidated as a description of the post-change bucket.**
  Under one consistent classifier of my own, the shipped 50 `not_found` ops split
  `25 object present but the keyword rule missed it / 13 no object at all / 12 kind mismatch`;
  the proposed 26 split **`13 under-action or paraphrase / 12 kind mismatch / 1 assignment
  starvation`**. So **kind mismatch goes from ~24% to 46% of the dominant bucket and becomes
  co-dominant**, because the proposal widens identity and leaves the kind filter untouched.
  The register's own phase-2 entry condition (changelog 2026-08-19: *"G-1, G-2, G-7 and the kind
  filter are one contract"*) is therefore **not met**: the proposal addresses G-1, partially
  addresses G-2, makes G-7 worse, and does not touch the kind filter.
- **V-3's 64/167 floor — unchanged and re-derived** (56 no-op emails + 8 all-cancel emails).
- **G-8** (`match` defaulting to the whole name) — genuinely dissolved: the contract never reads
  `op.match`. **C-10's goal of dropping the generated `match` field is achieved for the grader**,
  and blocked only by `sb/oracle.py:52` and `sb/schema.py:551-572` lint #5, which still consume it.

---

## 7. What I could not check

1. **Whether any of this is closer to the truth.** Phase 1d (the hand-grade) is deferred behind
   C-10, so there is no human reference. Every number here is grader-versus-grader; false-positive
   and false-negative rates against a human judgement remain unknown, which is exactly the gap the
   register says phase 2 must close. My §5 audit is a substitute, not a replacement: I judged the
   flips myself, from the capture.
2. **Generalisation.** One capture, one model, one seed (42), one lever set (1/5/7), one corpus
   sha. K-2 records 19 of 100 seeds raising `InfeasibleSchedule`, so seed-variance cannot be run.
   The +15 could be sampling noise of the same size as the 91→97 churn the register attributes to
   nondeterminism.
3. **Whether a real model would exploit any of this.** All adversaries are synthetic. No live run
   was made or requested (forbidden, and it would spend quota).
4. **Predicate coverage.** The corpus contains **0 `in`/`not_in` predicates, 1 `any_of`, and
   `exact_day` tolerance on 134/134 ops**, so the contract's behaviour under interval predicates
   and `within:Nd` tolerance is untested by any number in this document or in the proposal's.
   A `within:3d` op would make the `+3 days` adversarial pass outright.
5. **Downstream consumers.** I verified `sb/oracle.py`, `sb/scale.py`'s gate, and
   `sb/tests/test_e2e.py`. I did **not** check `sb/analyze.py`'s re-parsing of the grader's reason
   strings (`analyze.py:25`), `sb/schema.py` lint #5, or the vendored webapp copy under
   `webapp/api/_lib/sb/`.
6. **The stop list itself.** 100+ hand-written words. I used it verbatim as part of the contract
   and did not attempt to derive a principled version, nor to measure sensitivity to removing any
   single word. `list`, `final`, `event`, `day`, `review`, `meeting` and `budget`-adjacent terms
   are all stopped, and §2.1(b) shows `list` alone costs the reference oracle two emails.
7. **C-10 is still open.** The contract keys identity on `op.name`. If the corpus-authority
   question resolves toward production, several op names change, and every number here moves.

---

## Verdict table

| # | claim | verdict |
|---|---|---|
| 1 | shipped scores 97 / 148 / 64 / 64 | **CONFIRMED** (independent harness) |
| 2 | proposed scores 112 / 167 / 64 / 64 | **CONFIRMED** (independent harness; prototype agrees) |
| 3 | failure profile 50→26, 14→30, 11→0, 3→4, 2→2, 80→62 | **CONFIRMED**, every cell |
| 4 | "the +15 emails" | **OVERSTATED** — 16 up, 1 down |
| 5 | the +15 is genuinely earned | **CONFIRMED** — 0/16 flips for a bad reason; 0/69 passes survive ablation |
| 6 | `count: too many` 11→0 is real elimination | **CONFIRMED** — 7→ok, 4→wrong_day, none reclassified |
| 7 | identity from `op.name` removes the dependence on `match` | **CONFIRMED for the grader**; **WRONG end-to-end** — `sb/oracle.py:52` still writes `match` and drops to 161/167 |
| 8 | oracle = 167 shows the contract is honest | **UNSAFE** — circular; the repo's real oracle scores **161/167** and `sb/scale.py:127`'s 100% gate goes red |
| 9 | null = 64 shows no inflation | **CONFIRMED but vacuous** — structurally invariant across every variant tested |
| 10 | adversarial ≤ null shows no inflation | **UNSAFE** — a date-blind shotgun scores 93–148 (shipped 56); duplicates score 167 (shipped 64) |
| 11 | word-level matching is more robust than substring | **WRONG** — an inflecting perfect agent scores 68 vs the shipped grader's 127 |
| 12 | `cancel` unchanged in spirit | **UNSAFE** — widened to any-word; +1 false negative on the capture, 3 of the oracle's 6 new failures |
| 13 | attribution tie-break is safe | **CONFIRMED** as designed, **UNSAFE** as the `dateok` third term |
| 14 | assignment is non-deterministic / order-dependent | **OVERSTATED** — 40 shuffles, no flip; total sort key; list order only |
| 15 | greedy is worse than optimal | **OVERSTATED** — the one real divergence favours greedy; optimal fabricates a pass |
| 16 | single-word obligations are exploitable | **OVERSTATED at current volume** (0/69 ablation), **UNSAFE under volume** |
| 17 | the change satisfies the register's phase-2 entry condition | **WRONG** — the kind filter is untouched and becomes 46% of the residual bucket |

---

## Recommendation for the implementation, if it proceeds

Not a decision for a verifier, but these follow directly from the numbers above.

1. **Drop the `dateok` term from the tie-break.** Free on every measured world (real 112,
   oracle 167, `sb/oracle.py` 162, null 64, +3d 64) and removes 63 points of shotgun headroom.
2. **Put a volume brake back.** Not the old `count_ok` (G-2 shows it fires on distinct
   obligations, not duplicates), but something scoped to the *assignment*: e.g. fail an
   obligation if two or more objects tie at the top overlap **and** are attributed to the same
   email, or count unclaimed same-node objects against the no-action budget.
3. **Change `sb/oracle.py` in the same change-set**, or `sb.scale`'s mandatory gate fails.
   Per `CLAUDE.md`, this must not ride in the same commit as a corpus edit.
4. **Restore stemming or prefix tolerance** at the word level, or the contract is strictly less
   robust than the substring rule it replaces (claim 11).
5. **Do not widen `cancel` to any-word.** Grade it through the same assignment as create/move,
   which is what G-7's option list already recommends.
6. **Re-baseline `sb/tests/test_e2e.py:58`** deliberately and in writing, since it encodes the
   duplicate rule the proposal removes.
7. **Replace the guard set.** `+3 days` cannot separate any two contracts on this corpus. A
   guard set that would have caught all of the above: `dup×N`, `date-blind shotgun(span)`,
   `wrong-kind`, `inflected-title oracle`, `subject-title oracle`, and `sb/oracle.py` itself.
   All six are free and run in seconds against the capture.
8. Per the register's own status legend, this can reach **`fix proposed`** — not `applied`, and
   certainly not `verified`, since the phase-1d human reference does not exist.
