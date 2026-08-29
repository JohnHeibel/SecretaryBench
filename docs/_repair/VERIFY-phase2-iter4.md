# VERIFY — red-team of the phase-2 identity contract, iteration 4 (`grade_batch_v5`)

Third adversarial pass, over the **proposed-and-staged, not-yet-committed** iteration 4 of
the grader contract (`docs/grader-contract.md`, prototype
`docs/_repair/phase2_guards.py::grade_batch_v5`, staged port in the scratch overlay).
**This file reports; it changes nothing.** No edit to `sb/`, `corpus/`, `captures/`,
`webapp/`, the overlay, the prototype, or any other doc. No git operation. No live model
run. `scripts/recover_corpus.py` and `scripts/fix_match.py` were neither read nor executed.

**Method.** I wrote my own scoring harness from the prose of `docs/grader-contract.md`
rather than calling `grade_batch_v5` or the overlay's `grade_batch`, so a bug in the
prototype's measurement code could not hide. `ref4.py` re-implements tokenisation, the
stop list, the new stemmer, keyword extraction, title/full overlap, title precision, the
eight-term ranking, exclusive greedy assignment, the create-freshness rule, the day-scoped
volume brake, the stale rule, the joint cancel phase, the wrong-kind report, the no-action
branch, both grading paths and the synthetic agents. It borrows from `sb/` only plumbing
that is identical under every contract: `resolver`, `schema.load_corpus`,
`scheduler.build_plan`, `live/runner._node_state` / `_turn_delta`, `oracle._target` /
`_as_dt`, and `grader._predicate_ok` / `_fmt_obj` / `_describe_predicate`. I then ran the
prototype's own `report5([...])` and diffed the overlay's `grade_batch` against mine at
**verdict *and* per-op reason *and* per-op `actual` level** on 21 worlds × 2 paths.
Scratch code lives outside the repo at
`/private/tmp/claude-501/-Users-jamesoc-dev-SecretaryBench/cf81dddf-4a9b-49d4-996d-ea65d999c189/scratchpad/verify4/`
— `ref4.py` (contract), `ref4_worlds.py` (worlds), `adv.py` / `adv2.py` (new adversaries),
`cases.py` (constructed cases), `profile.py`, `look.py`, `dump_ref.py` / `dump_overlay.py`
(the port diff), `mutate.py` (guard mutation testing).

---

## Bottom line

**Yes — safe to implement, with four changes, all of which I measured to cost zero.**

This is the first iteration whose numbers I could not move. Every cell of the measured
table, the op-level profile, the "0 flips" claim, the stop-word leave-one-out and the
oracle gates reproduce **to the digit** under an independent implementation; the staged
overlay is byte-identical to the prototype on every world and every reason string; 60 pool
shuffles × 21 worlds × both paths produce **0 verdict flips**; `sb.scale` gates at 167/167
and 287/287; the 104-test suite passes. The two properties that killed iteration 3 —
laundering and order-dependence — are genuinely closed on the measured evidence.

It is not a clean pass, because **three safety properties are stated as facts in the
document and are false**, and because **the document credits the wrong mechanisms for two
of its own fixes**. None of the three costs a point on any measured world, which is why the
guard set does not see them; all three are one-line fixes.

1. **Rule 3's "That is a duplicate, whatever id it carries"
   (`docs/grader-contract.md:95`) is FALSE.** Duplicates stamped with an email id from
   **another node**, or with an id that resolves to nothing, are dropped from the node pool
   by `sb/live/runner.py:139-145` before the grader ever sees them. A scheduling-perfect
   agent that leaves 5 such copies of every object scores **167/167 with 594 objects in the
   store where the perfect agent holds 99** — and 2079 objects at 20 copies, for the same
   167. `sb/live/runner.py:107` tells the model *"Never leave duplicates."* This is the
   iteration-3 exploit relocated from the email to the node. It is **not a regression**
   (the shipped grader is equally blind: 158, its own ceiling), and it is A-5, which the
   contract explicitly declines to fix — but the document asserts the opposite in the
   contract text and the changelog says laundering is "guarded by three launder worlds".
   All three launder worlds stamp within the node.

2. **"No attribution rule … Nothing passes or fails on it"
   (`docs/grader-contract.md:121-122`) is FALSE.** Rank step 4 (`created for this email`,
   overlay `sb/grader.py:300`) sits **before** rank step 6 (`create/move before cancel`,
   `:302`), so the stamp, not the verb, resolves a create-vs-cancel tie. Constructed: one
   node, one day, `create 'Vendor Onsite'` and `cancel 'Vendor Onsite'`, the model does
   exactly the right thing — stamping the new object with the **cancel's** email id instead
   of the create's flips **both emails from pass to fail**. Step 4 is measurably **inert**:
   removing it changes 0 verdicts and 0 scores on all 21 worlds. It is pure downside.

3. **"Steps 7–8 make the ranking a total order on content: the verdict cannot depend on the
   order the store lists objects in" (`docs/grader-contract.md:89`) is FALSE as a
   property**, though true as a measurement. The final key is
   `(…, ei, oi, o.title, o.when.isoformat())` — **description is not in it**. Two same-kind
   objects with the same title, time and stamp but different descriptions have *identical*
   sort keys, so `list.sort`'s stability hands the choice to pool order. Constructed: two
   events "Team Offsite" on the same day, one with description "budget review"; ops
   `create 'Team Offsite'` + `create 'Budget'` score **2 passes in one pool order and 0 in
   the other**.

**And two mechanisms are misattributed, which matters because the register will record
them:**

- **Batch grading is inert.** Grading each email alone against a **day-scoped** `today`
  gives byte-identical results on all 21 worlds and the real capture, except **one op's
  reason string** (a stale-copy count, 2 vs 1). Laundering scores 64/64/77 either way. The
  anti-laundering fix is the **day scope**, exactly what `VERIFY-phase2-iter3` §12.1
  recommended and measured; the batching is not doing the work the document credits it
  with. Worse, batching is what *created* the `manufacturing-kickoff-3` exposure the
  document then congratulates itself for closing.
- **Joint cancel ranking is inert, and it is not what closes `manufacturing-kickoff-3`.**
  Running cancels in a second phase (iteration 3's order) produces **0 per-op differences
  across all 21 worlds**. What closes the false pass is **title precision**: with
  `precision` disabled the cancel is shielded and `real` reads **115**. The document's
  stated reason ("the cancel matches the walk-through on four words, the move on one") is
  not how the ranking gets there — title overlap is the *first* key and the move scores
  1.00 on that object against the cancel's 0.25, so the move outranks the cancel and is
  only steered elsewhere by precision (0.33 on `"WC cleat reveal event"` vs 0.20 on the
  walk-through).

Smaller, but each is a factual error in a document that will become the register's record:
**5 of the 6 `wrong kind` labels are genuine, not 6 of 6**; **7 `wrong day` diagnoses
became `not found`, not 8** (the eighth became `wrong kind`); **`event day!` is not one of
the 6 wrong-kind labels** — it is one of the three that stopped being one; the
`wrongkind` +3 audit describes the wrong mechanism; and the `dupmove_retitle` guard's
computed threshold is **10 when 11 are caught**, leaving a point of slack.

| | shipped | iter. 3 | **iter. 4** | null |
|---|---|---|---|---|
| real (certified 1c capture) | 97 | 114 | **114** | 64 |
| `oracle_engine` / `oracle_name` | 158 / 158 | 167 / 167 | **167 / 167** | — |
| every doc-table cell, 18 columns | ✓ | ✓ | **✓ reproduces exactly** | — |
| 60 pool shuffles × 21 worlds × 2 paths | 0 flips | flips | **0 flips** | — |
| **NEW `launder_cross` — 5 copies stamped with *another node's* email id** | 158 | — | **167** ✗ (594 objects) | 64 |
| **NEW `launder_bogus` — 5 copies stamped `no.such.email`** | 158 | — | **167** ✗ (594 objects) | 64 |
| **NEW `dupkind` — 5 copies of every object, of the *opposite* kind** | 153 | — | **166** (2079 objects at ×20) | 64 |
| **NEW `onestamp` — one turn, everything stamped with the first email's id** | 82 / 96 | — | **83 runner / 97 engine** | 64 |
| **NEW `oneword` — perfect dates, title = ONE content word of the name** | 105 | — | **165** | 64 |
| **NEW `defer1` — perfect, then 5 copies of everything the next day** | 140 | — | 144 | 64 |
| constructed: stamp a correct object with a sibling **cancel's** id | pass/fail | — | **both emails fail** | — |
| constructed: two same-titled objects, different descriptions | — | — | **2 pass or 0 pass, by pool order** | — |
| constructed: two cancel emails, one survivor carrying both names | both fail | — | **one email passes falsely** | — |
| constructed: object made in a sibling's turn earlier the same day | pass | — | **runner pass / engine fail** | — |

`launder_cross` meets the brief's regression criterion in its second form: it scores 167
while violating the one bookkeeping instruction the prompt states, and no guard sees it.

---

## 1. Reproducing the measured table — CONFIRMED, every cell

`ref4.py` (mine) against `phase2_guards.report5([...])` (the prototype) against
`docs/grader-contract.md:174-192`. All three agree on **every cell of every row**.

| world | shipped | iter. 3 | **iter. 4 — doc / prototype / mine** |
|---|---|---|---|
| real | 97 | 114 | **114 / 114 / 114** |
| `oracle_engine` | 158 | 167 | **167 / 167 / 167** |
| `oracle_name` | 158 | 167 | **167 / 167 / 167** |
| `oracle_subject` | 95 | 137 | **137 / 137 / 137** |
| `oracle_inflect` | 142 | 159 | **165 / 165 / 165** |
| `null` · `dup5` · `shot7` · `shot45` · `shot90` | 64 | 64 | **64 (all five)** |
| `wrongkind` | 65 | 69 | **67 / 67 / 67** |
| `dupmove` | 148 | 152 | **152 / 152 / 152** |
| `dupmove_retitle` | 146 | 163 | **156 / 156 / 156** |
| `launder` | 72 | 167 | **77 / 77 / 77** |
| `launder_all` | 64 | 152 | **64 / 64 / 64** |
| `launder_past` | 64 | 139 | **64 / 64 / 64** |
| `bothkinds` | 153 | 166 | **166 / 166 / 166** |
| `nocancel` | 150 | 159 | **159 / 159 / 159** |

The prototype's own checker prints `iter4  PASS` and correctly prints
`shipped FAIL -> oracle_engine 158; oracle_name 158; nocancel 150; dupmove 148` — the
`dupmove ✓` error that `VERIFY-phase2-iter3` §1 found in the iteration-3 document is
**fixed**; the guard table now marks it `148 ✗`.

**Op-level failure profile — CONFIRMED, every cell** (190 details = 134 ops + 56
no-action):

| bucket | shipped (doc / mine) | iter. 3 (doc / mine) | iter. 4 (doc / mine) |
|---|---|---|---|
| ok | 56 / 56 | 76 / 76 | **76 / 76** |
| not found | 50 / 50 | 18 / 18 | **28 / 28** |
| wrong kind | — | 8 / 8 | **6 / 6** |
| wrong day | 14 / 14 | 29 / 29 | **21 / 21** |
| count: too many | 11 / 11 | 0 / 0 | **0 / 0** |
| stale copy on move | — | 1 / 1 | **1 / 1** |
| cancel residue | 3 / 3 | 2 / 2 | **2 / 2** |
| over-acted | 2 / 2 | 2 / 2 | **2 / 2** |
| **failing details** | **80 / 80** | **60 / 60** | **60 / 60** |

(As in iteration 3, the `ok` row excludes the 54 passing no-action details; raw is 130.)

**"0 flips against iteration 3" — CONFIRMED.** Both 114; the passing sets are identical.
There are **14 per-op reason changes**, and their composition is *not* what the document
says (§8 below).

**Other numbers, all CONFIRMED to the digit:**

- Stop list has **97** words; **66** never occur in an op name; leave-one-out over all 97
  moves some world for exactly **7** words (`call`, `day`, `final`, `list`, `meeting`,
  `planning`, `review`), never by more than **±2**, and **none moves `oracle_name`**.
- **24 of 134** ops have a single content word.
- The shipped grader's 9 losses on the name-titled perfect agent: **2** whole-name
  substrings (`Team_pizza_party`, G-8) + **7** `found 2 matching, expected exactly 1` (G-2),
  **0** agent faults.
- `dupmove` = 167 − 15 and the 15 failing emails are **exactly** the 15 move emails.
- `nocancel` = 167 − 8; all 8 all-cancel emails fail, 0 false-passing cancel ops.
- `stale_floor ≥ half` breaks both oracles at **166 / 166** (open issue 2).
- `stop_on_stem=True` costs **oracle_subject 1** and nothing else (the doc's stated
  measurement).
- 0 same-node same-kind op names are substrings of a sibling name, so `Store.find_in_node`
  stays unambiguous under `op.name` titles.
- **`sb.scale` on the overlay: `oracle: 167/167 = 100%` (`--filler 0`) and
  `287/287 = 100%` (default).** Both gates green.

---

## 2. The staged port is the prototype — CONFIRMED, at reason level

`dump_ref.py` vs `dump_overlay.py`, comparing `(passed, [(reason, actual) per detail])` for
every email:

```
21 worlds × {runner, engine} = 41 comparisons
verdict differences : 0
detail differences  : 0     (reason AND actual string, every op, every world)
```

That includes `real` (114/114), all four oracles, all five volume worlds, all three
launders, `dupmove`/`dupmove_retitle`, `bothkinds`, `nocancel`, and the four new
adversaries I added. `oracle_engine` through the *real* `sb.engine.run` with the overlay's
`sb/oracle.py`: **167**.

**Wiring — CONFIRMED.**

- `overlay/sb/live/runner.py:161-187` defines `_grade_day`; the day loop calls it at
  `:637` and reads `graded[eid]` at `:641`. There is no second `grade_email` call site left
  in the runner.
- `overlay/sb/regrade.py:60` and `overlay/sb/tests/test_capture_regrade.py:80` both route
  through the same `_grade_day`. `python -m sb.regrade captures/baseline-sonnet-4-5` with
  the overlay first on `sys.path` prints **`RESCORE 114/167`**.
  *(Note for whoever reproduces this: `PYTHONPATH=<overlay> python -m sb.regrade …` run from
  the repo root silently uses the **shipped** grader, because `sys.path[0]` is the cwd. It
  prints 97. `sys.path.insert(0, overlay)` gives 114. The `pytest` invocation in the brief is
  fine — pytest inserts the package basedir first.)*
- `_grade_day` recomputes `by_eid` from `state` rather than reading the capture's recorded
  `by_eid`. I checked all 57 day records: **0** of the `day_new` ids disagree, so the
  substitution is behaviour-preserving on this capture.
- `sb/engine.py` is unchanged and reaches the same code through
  `grade_email → grade_batch([one], turn)`. On the 21 worlds the engine and runner paths
  agree everywhere except the two mis-stamp worlds (§5.4).
- **104 tests pass in ~1.7 s** in the overlay; 65 pass on HEAD.

---

## 3. Laundering — closed inside the node, wide open across nodes — UNSAFE AS STATED

### 3.1 The claim

> A `create`/`move` **fails** if another same-kind object **created today** matches it at
> least as well and was claimed by no op in the batch. **That is a duplicate, whatever id it
> carries.** — `docs/grader-contract.md:93-95`

and, in the changelog, *"laundering scores 64 / 64 / 77 … both paths agree on all 17
worlds"*, guarded by `launder` / `launder_all` / `launder_past`.

All three launder worlds stamp copies with **an email of the same node**
(`phase2_guards.py:741-746`, overlay `test_grader_guards.py:114-120`). The pool the brake
scans is `items[0][2]`, the **node** state, built by `sb/live/runner.py:139-145`:

```python
e = corpus.emails.get(o.email_id)
return o if (e and e.node == node) else None
```

An object stamped with another node's email id, or with an id that resolves to nothing, is
therefore **absent from the pool** — not unclaimed, absent. `today` (`_grade_day` line 173)
still contains it, but the brake iterates `pool`, never `today`.

### 3.2 Measured

```
world                                    objects   shipped   iter. 4   null
perfect control                              99       158       167      64
launder      (5 copies, node sibling id)    509        72        77      64
launder_cross  (copies -> another node)     594       158       167      64
launder_bogus  (copies -> 'no.such.email')  594       158       167      64
launder_cross,  1 copy / 20 copies      198 / 2079    158       167      64
```

Volume-independent, path-independent (engine 167 = runner 167), shuffle-invariant
(60 shuffles, 0 flips). The agent schedules perfectly and leaves five copies of every
object on the user's calendar, and the grader awards full marks.

Two refinements, for the record:

- A cross-node stamp *does* trip `store_app.py:86-92` (`invalid_email_id` or
  `stale_email_id`). Warnings are monitor-only and **nothing aggregates them** (register
  A-4), so this is not a defence.
- A **warning-free** version exists: stamp the copies with an email of another node served
  **today** — `_watch_attribution` has no branch for that (A-4's blind spot), and A-5
  records that 150 of 167 emails have a same-day sibling in a different node. Measured, the
  naive version of this costs the launderer points because the copies pollute the target
  node (`146` runner / `167` engine, and shipped splits the same way at 137/158); a
  launderer that picks a target node with no later emails is warning-free and free, but only
  19 of 125 ops can do so on this plan (164/167).

### 3.3 Verdict

Not a regression against the shipped grader, and squarely inside A-5, which the contract
declines to fix. But rule 3's sentence is false as written, and the guard set's coverage is
described as complete when it covers only same-node stamps. **The register must not record
A-5 as unaffected by this contract a second time** — `VERIFY-phase2-iter3` §2.3 made exactly
that request.

---

## 4. Batch grading is inert; the day scope is the fix — MISATTRIBUTED

### 4.1 Batching changes nothing

Grading each email **alone** against a day-scoped `today` (i.e. keeping every other rule,
dropping only the joint assignment):

| | real | oracle_name | oracle_subject | oracle_inflect | null | dup5 | shot* | wrongkind | dupmove | dupmove_retitle | launder | launder_all | launder_past | bothkinds | nocancel |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| batched (as designed) | 114 | 167 | 137 | 165 | 64 | 64 | 64 | 67 | 152 | 156 | 77 | 64 | 64 | 166 | 159 |
| **per-email, day `today`** | **114** | **167** | **137** | **165** | **64** | **64** | **64** | **67** | **152** | **156** | **77** | **64** | **64** | **166** | **159** |

Across all 21 worlds and the capture: **1 per-op reason difference, 0 verdict differences.**
The one difference is cosmetic —
`World_Cup_Cleat_Launch.reveal-event-date-and-venue-3`, `move 'Reveal Event'`, reads
`1 stale copy` batched and `2 stale copies` per-email, because batching lets the sibling
email's cancel claim the walk-through.

The measured anti-laundering benefit is entirely the **day scope**. Re-scoping the brake
back to the per-email turn (`brake_scope='mine'`) reproduces iteration 3's numbers exactly:
`launder 167`, `launder_all 152`, `launder_past 139`. So the document's §2 row credits
`grade_batch` for what `today` does.

### 4.2 Batching *created* the exposure it is credited with closing

`docs/grader-contract.md:48` says batch grading revealed a real-capture shielding case that
joint cancel ranking then closed. Measured:

```
                              kickoff-3 cancel     real
per-email, precision on           FAIL             114
per-email, precision OFF          FAIL             114
batched,   precision on           FAIL             114
batched,   precision OFF          PASS (false)     115
```

Under per-email grading the cancel cannot be shielded at all, with or without precision.
Under batching it can, and the only thing preventing it is the **title-precision tie-break**
— 0.33 for `"WC cleat reveal event"` against 0.20 for `"Design walk-through at WC reveal"`,
both at title overlap 1.00 and full overlap 1.00 for the single keyword `{reveal}`.

### 4.3 Joint cancel ranking is inert

Running cancels as a second phase (iteration 3's ordering) inside the batch:
**0 per-op differences across all 21 worlds**, `real` 114, `kickoff-3` still fails. The
document's mechanism — *"the op that matches more of an object's words gets it (the cancel
matches the walk-through on four words, the move on one)"* — is not what happens: **title
overlap is the first key**, and on that object the move scores 1.00 (`reveal` is in the
title) against the cancel's 0.25 (only `design` is), so the move sorts first regardless of
words matched. Removing the `words matched` term entirely leaves the whole 104-test suite
green (§10).

---

## 5. The ranking

### 5.1 Order-invariance — CONFIRMED as a measurement, far beyond what was tested

60 pool shuffles × 21 worlds × both grading paths (the doc claims 8 × 17, runner only):

```
every world: canonical score in 60/60 shuffles, 0 per-email verdict flips, on both paths
real 114 · oracle_name 167 · oracle_subject 137 · oracle_inflect 165 · dupmove 152
dupmove_retitle 156 · launder 77 · launder_all 64 · launder_past 64 · bothkinds 166
nocancel 159 · wrongkind 67 · null/dup5/shot* 64 · launder_cross 167 · onestamp 83/97
```

Iteration 3's five order-sensitive worlds (`oracle_name` 165–167, `oracle_subject` 135–136,
`oracle_inflect` 157–159, `dupmove` 151–152, `nocancel` 158–159) are all pinned. The
`sb.scale` gate is deterministic. **This is a real and complete fix of iteration 3 §3.**

### 5.2 …but the key is not a total order — the property claim is WRONG

`overlay/sb/grader.py:297-304` ends the key with `ei, oi, o.title, o.when.isoformat()`.
`o.description` and `o.email_id` are absent. Two same-kind objects with equal title, time
and stamp therefore produce **identical keys**, and `list.sort` stability defers to pool
order. Constructed (`cases.py` CASE 4):

```
pool: event "Team Offsite" Jun 12  desc "budget review"
      event "Team Offsite" Jun 12  desc ""            (same stamp, same turn)
ops : create 'Team Offsite', create 'Budget'
   pool [x, y] -> over-created: 2 equally-matching events ; no event titled like "Budget"   (0/2 ops)
   pool [y, x] -> matched ; matched                                                          (2/2 ops)
```

The store accepts exact-duplicate creates (`store_app.py:127-131`, no dedupe), so a model
that writes two same-titled entries and describes only one reaches this. Appending
`(o.description, o.email_id)` to the key closes it and costs nothing (no world has such a
pair, which is why 60 shuffles pass).

### 5.3 Rank step 4 (`mine`) makes a stamp decide a verdict — UNSAFE, and free to remove

`docs/grader-contract.md:121-122`: *"**No attribution rule.** `email_id` decides only which
email's *turn* an object belongs to (the no-action check and rank step 4). Nothing passes or
fails on it."* Step 4 is at `overlay/sb/grader.py:300`, **two positions before** the verb
priority at `:302`. Constructed (`cases.py` CASE 1), one node, one day:

```
ops: email A = create 'Vendor Onsite' (event) ; email B = cancel 'Vendor Onsite' (event)
model: deletes the old object, creates the new one on the right day  -- correct behaviour

  new object stamped A (correct)   -> A matched      · B cancelled       2/2 PASS
  new object stamped B (sibling)   -> A "not found"  · B "still on the calendar"  0/2 FAIL
  new object stamped C (a third sibling) ->                              2/2 PASS
```

The middle row is the register's A-4 blind spot exactly: a sibling id from today's batch,
which the store does not warn about. Two emails lost for a transcription error, in a
contract that says nothing passes or fails on the stamp.

**And step 4 earns nothing.** Removing it: identical scores on all 21 worlds, **0 verdict
differences**, 16 per-op `actual`-string differences (which object is named), 0 reason
differences. It is pure risk. Moving it *after* the verb term would keep whatever
diagnostic value it has and remove the create-vs-cancel exposure.

### 5.4 Path agreement — holds where it was tested, and `create_fresh_only` opens a new class

The 19 correctly-stamped worlds agree exactly between paths. The two mis-stamp worlds do
not, and **the shipped grader splits by the same amount**, so this is A-2, not a regression:

```
onestamp (one turn, everything stamped with the first email's id)
   shipped   runner 82  / engine 96
   iter. 4   runner 83  / engine 97
```

More interesting is a class iteration 4 **introduces**. A `create` may claim only an object
in `today`; `today` is the whole day on the runner path and *this email's own model call* on
the engine path. Constructed (`cases.py` CASE 8), identical state:

```
pool: event "Board Sync" Jun 12, stamped n.b ; op: create 'Board Sync' in email n.a
   runner  (today = the day)        -> matched   PASS
   engine  (today = own turn only)  -> "no event titled like Board Sync was created"  FAIL
```

The shipped grader passes both (it has no freshness rule). The measured worlds never
exercise it because each synthetic email creates its own objects in its own turn. Not a
blocker — but the document's "both paths agree" should be scoped to "on the worlds
measured", and this is the shape to add if anyone wants it guarded.

---

## 6. The stale rule and the `dupmove_retitle` arithmetic

### 6.1 The floor change works, and it is free on the oracles — CONFIRMED

`> half the words` catches **11 of 15** retitled double-bookings (iteration 3's `≥ claimed
score` caught 4). Dropping the *first* word instead of the last gives the same 156, so the
evasion is not last-word-specific. `no stale rule` gives `dupmove 167` and
`dupmove_retitle 167`, so the rule is what does the catching. Removing it costs
`oracle_subject` **nothing** (137 both ways) — iteration 3 lost 1 point here, and the
document's implicit claim that this is now free **CONFIRMED**.

### 6.2 The guard's threshold is computed wrong — 10 vs 11

`test_grader_guards.py:52-54` (and `phase2_guards.py:838-840`) compute

```python
RETITLE_CAUGHT = sum(1 for e in EMAILS.values()
                     if any(op.verb == "move" for op in e.answer.ops)
                     and not all(len(keywords_of(op)) == 2 for op in e.answer.ops if op.verb == "move"))
```

which assumes the dropped last **word** is a **keyword**. It is not when the name ends in a
stop-word. `World_Cup_Cleat_Launch.endorsement-terms-for-the-reveal-2`, obligation
`'Endorsement LOI Review'`: `review` is a stop word, so `keywords_of` = `{endorsement, loi}`
(len 2 → predicted to escape), but dropping `Review` leaves both keywords → 1.0 > 0.5 →
**caught**. Computed 10, actually caught 11, so the guard asserts `≤ 157` against an actual
156: **one point of slack**, i.e. one more double-booking could slip through and the guard
would still say PASS. Compute the threshold by simulating the drop, not by counting
keywords. (The 4 that genuinely escape are all two-keyword names, as the doc says:
`Company Retreat`/`Contact Retreat Location`, `board_signoff`, `Press Briefing`,
`expo_keynote`.)

`test_laundered_duplicates_on_creates_keep_only_the_move_emails` (`:203`) has the same
shape of slack: bound `FLOOR + MOVE_EMAILS = 79`, actual 77.

### 6.3 The stale rule still fires only on a description — CONFIRMED, and it is stricter now

Its one firing on the capture is `move 'Reveal Event'` (`{reveal}` — `event` is a stop
word) flagging `"WC cleat press briefing"` whose **description** reads *"…week before
reveal"* (title overlap 0.00, full overlap 1.00). True-positive rate on the capture:
**0 of 1**, as the document says. And because the floor is now an absolute `> 0.5` rather
than `≥ the claimed score`, a move that claims a weakly-titled object can be failed by a
better description-only match:

```
op   : move 'Vendor Kickoff'  ({vendor, kickoff})
claim: event "Vendor sync"       (title 1 of 2 -> 0.5, wins on title overlap)
extra: event "Quarterly planning" desc "vendor onboarding kickoff"  (title 0.0, full 1.0)
   -> "moved, but 1 stale copy left behind (double-booked)"   FAIL
```

This is G-3 acting against the model. It is the price of the stricter floor and it should be
stated next to the 11-of-15 improvement.

---

## 7. The cancel rule

### 7.1 τ = 1.0 and same-kind claiming — CONFIRMED, including the historical G-7 case

`nocancel` = 167 − 8 exactly, 0 false-passing cancel ops. The canonical G-7 failure — the
conflated `"Boston tech trip (WHOOP + Boston Dynamics)"` that failed `~"dynamics"` in **all
four** recorded runs — now **passes** (overlap 0.67 < 1.0), while a genuinely undeleted
`"Boston trip - Boston Dynamics visit"` still fails. Cross-kind survivors are correctly out
of scope (the shipped semantics), closing `VERIFY-phase2-iter3` §5.4. Both remaining capture
failures are genuine and unchanged.

### 7.2 Exclusive claiming makes `cancel` non-monotone — a new false-pass shape

A survivor can only accuse **one** cancel. Constructed (`cases.py` CASE 5), from the shape
the corpus already contains (`shoe-product-launch-delays.re-design-team-meeting` carries
`cancel 'Decide who to cut'` `{cut, who}` and `cancel 'Design cut meeting'` `{cut, design}`):

```
survivor: event "Design cut meeting - who to cut"   (carries every word of BOTH)
 both cancels in ONE email  -> op0 fails, op1 reads "cancelled"  (email fails anyway)
 the two cancels in TWO emails of one node, one day
                            -> email0 FAILS, email1 PASSES  <-- false pass
 shipped grader, either email alone -> FAILS
```

The model deleted nothing and one email is marked correct. **Latent on this corpus** — each
of the 15 nodes has exactly one cancel email, so two cancel emails never share a node-day —
but it is a phase-5 corpus edit away, and it is a *regression against the shipped rule*,
which evaluates each cancel independently. The in-email version is live today and produces a
false `cancelled` label in the log.

### 7.3 The sibling-shielding surface is now the other way round

`VERIFY-phase2-iter3` §5.3's constructed false pass (a create claiming the object a
same-email cancel should have caught) is **closed** by `create_fresh_only`: with the model
doing nothing, the create reports `not found` and the cancel reports `still on the calendar`
(`cases.py` CASE 2). Confirmed. What replaced it is §4.2 above — batching lets a *sibling
email's* create/move outrank a cancel, held off only by the precision tie-break.

---

## 8. The diagnosis work — smaller than claimed

### 8.1 "Eight `wrong day` diagnoses became `not found`" — WRONG, seven did

The 14 per-op reason changes from iteration 3 to iteration 4:

| change | n |
|---|---|
| `on the wrong day` → `no <kind> titled like …` | **7** |
| `on the wrong day` → `wrong kind` | **1** (`Sponsoring-Marathon.approval-of-budget-tier`) |
| `wrong kind` → `no <kind> titled like …` | 3 |
| `moved, but 2 stale copies` → `… 1 stale copy` | 1 |
| `over-acted` → `over-acted — created …` (restores the shipped wording) | 2 |

`wrong day` 29 → 21 is −8, but only 7 of them became `not found`.

### 8.2 "6 of 6 wrong-kind labels are genuine" — WRONG, 5 of 6

For each of the 6 labels I checked whether the reported object was created **today** and
stamped to **this** email — i.e. whether it really is this model's answer, of the wrong kind:

| email | reported object | ov | created today | stamped to this email | genuine |
|---|---|---|---|---|---|
| `pizza-party.end-of-year-pizza-party` | "End-of-year pizza party" | 0.67 | yes | yes | ✓ |
| `Innovation-comp.heads-up-one-of-the-pitches-…` | "Ask about patent conflict…" | 0.67 | yes | yes | ✓ |
| `Innovation-comp.one-of-our-designers-got-a-job-offer` | "Retention conversation…" | 1.00 | yes | yes | ✓ |
| `Innovation-comp.found-a-typo-on-the-trophy` | "Approve corrected trophy…" | 0.67 | yes | yes | ✓ |
| `press-tour.podcast-taping-pick-a-day` | "Pick podcast taping date" | 1.00 | yes | yes | ✓ |
| **`Sponsoring-Marathon.approval-of-budget-tier`** | **"Select Eugene Marathon sponsorship tier"** | **0.67** | **no** | **no — stamped `Sponsoring-Marathon.sponsorship-tiers`** | **✗** |

The sixth is an **earlier day's answer to a different obligation**, matched on
`sponsorship` in the title plus `budget` in its description ("…ensure it aligns with
budget"). The model created nothing for `sponsorship & budget approval meeting`; the grader
prints `wrong kind: created a to-do, expected a event`. This is precisely the class
`VERIFY-phase2-iter3` §6.4 flagged (2 of 8 false) and the document claims to have removed.
The wrong-kind report excludes objects claimed by a sibling op **in this day's batch**; an
inherited object from an earlier day is never claimed and is therefore always eligible.

**One-line fix, measured free:** require the wrong-kind candidate to be in `today` (or
raise the floor to 0.75). Either removes the false label and keeps all five genuine ones;
score-neutral on every world (wrong-kind objects are reported, never claimed).

### 8.3 `event day!` is not one of the 6 — WRONG

`docs/grader-contract.md:213-215` lists the 6 as including *"the all-stop-word `event
day!`"*. `Sponsoring-Marathon.approval-of-event` is one of the **three that stopped** being
a wrong-kind label under iteration 4; it now reads
`no to-do titled like "event day!" was created`.

### 8.4 The `wrongkind` +3 audit describes the wrong mechanism — WRONG

`docs/grader-contract.md:216`: *"three emails where a right-kind sibling object shares the
obligation's single content word and sits on the right day."* The three emails are:

```
shoe-product-launch-delays.design-team-meeting
   create 'Design cut meeting' (event, {cut,design})  matched the swapped EVENT "Decide who to cut"
   create 'Decide who to cut'  (todo,  {cut,who})     matched the swapped TODO  "Design cut meeting"
Company-Retreat.planning-call-and-forms-for-your-company
   create 'Retreat Company Meeting Call' (event, {company,retreat}) matched EVENT "Fill out Retreat Forms" (ov 0.50)
   create 'Fill out Retreat Forms' (todo, {form,out,retreat})       matched TODO  "Retreat Company Meeting Call" (ov 0.33)
Partnership-with-deeptech-companies.boston-partnership-trip   (same shape)
```

Not a single-content-word obligation among them, and no "right-kind sibling": each email has
**one event op and one todo op with overlapping vocabulary**, and the `wrongkind` agent
swaps both kinds, so the two ops **claim each other's objects**. A model that gets *both*
kinds wrong on such an email scores full marks on it. That is a more interesting weakness
than the one described, and it is a direct consequence of §9.1.

---

## 9. Identity

### 9.1 There is no minimum overlap for a create/move claim — unstated

`overlay/sb/grader.py:295` admits any pair with `sc > 0`. Consequences, measured:

- On the real capture, of 69 passing create/move claims, **8 sit at exactly 0.5** and **3
  below** — 16% of the recovered passes rest on half the obligation's words or fewer
  (`'AI_Sign_Off'` claiming `"Sign Anthropic deal documents"` at 0.33;
  `"Double check Josh's advertising work"` claiming `"Review Josh's marketing
  advertisements"` at 0.40, twice).
- On `oracle_subject`, 35 of 82 passing claims are at ≤ 0.5.
- **New adversary `oneword`**: a scheduling-perfect agent that titles every object with a
  **single content word** of the obligation name scores **165/167** (shipped 105). A
  `halfname` agent (first half of the name) scores **157** (shipped 105).

This is a deliberate trade against G-1 and it is why the contract can recover +17. But the
document never states it, and it is the mechanism behind both `wrongkind` = 67 (§8.4) and
`oracle_subject` = 137. It belongs in "Explicitly not included" alongside "no `dateok`
term", so nobody later reads the +17 as evidence that identity became *precise*.

### 9.2 The stemmer — genuinely better, with two residues

Op-name words whose proper English plural still fails to unify: **36 of 207** under
iteration 4 against **59 of 207** under iteration 3. `note/notes`, `date/dates`,
`code/codes`, `trophy/trophies` now unify; `oracle_inflect` 159 → 165 reproduces.

Residues worth recording:

- **`freeze` / `freezes` split** (`Design Freeze Sign Off` is a live obligation): `freezes`
  ends in `es` after `z`, so it strips two characters to `freez`, while `freeze` is left
  whole. The sibilant rule needs a "singular already ends in `e`" case.
- **The stop list is still applied pre-stem, and the document knows it.** `news → new`
  (and `new` *is* a stop word), `meetings → meet` (and `meeting` is). Live: the cancel
  `press-tour.news-hit-fell-through` has keywords `{new, segment}`, so any same-kind object
  titled "New … segment" fails it; `Send CTO Rubric from Meetings` carries `meet`, matched
  by every title containing "meeting". The document discloses the trade-off (stop-on-stem
  costs `oracle_subject` 1) but not the residue.
- 13 distinct corpus/capture word pairs share a stem; all except `new/news` are desirable
  (`company/companies`, `note/notes`, `william/williams`, …).
- One-character keyword in the corpus: `"Double check Josh's advertising work"` yields
  `{advertis, double, josh, s, work}` — `s` matches any standalone "s" token.

### 9.3 The empty-name guard — CONFIRMED, and complete

`overlay/sb/schema.py:139-141` raises
`CorpusError: op create name '!!!' has no letters or digits …`. `re.search(r"[A-Za-z0-9]")`
is exactly the condition under which `keywords_of` returns the empty set, so the guard's
coverage is complete for the vacuous-cancel case `VERIFY-phase2-iter3` §7 raised.

---

## 10. The guard set, mutation-tested — two terms are unguarded

A guard that cannot fail is worth nothing, so I mutated the overlay's `sb/grader.py` one
rule at a time and ran the 104-test suite (`mutate.py`).

| mutation | tests failed | detected by |
|---|---|---|
| remove the duplicate rule | 9 | dup5, shot7/90, launder_all/past, launder, bothkinds, 2 unit |
| cancel τ 1.0 → 0.5 | 7 | both oracles, dupmove, 3 pins, real pin |
| stale floor `> half` → `≥ half` | 7 | both oracles, nocancel, 3 pins, real pin |
| brake scoped to the email's turn (iter-3 bug) | 6 | 2 launder pins, launder-creates, 2 path tests, 1 unit |
| cross-kind claiming | 6 | 3 pins, real pin, 2 unit |
| remove the stale rule | 4 | dupmove, dupmove_retitle, test_e2e, 1 unit |
| drop the `(title, when)` tail of the key | 2 | oracle_subject pin + its shuffle test |
| create may claim any object | 2 | wrongkind pin, 1 unit |
| remove `mine` (step 4) | 2 | test_e2e, 1 unit *(0 world scores move)* |
| remove title-overlap-first (step 1) | 1 | 1 unit only *(0 world scores move)* |
| remove title precision (step 5) | 1 | the `real` pin (114 → 115) |
| batch → per-email (day `today` kept) | 1 | 1 unit only |
| **remove `words matched` (step 3)** | **0** | **nothing** |
| **remove verb priority (step 6)** | **0** | **nothing** |
| wrong-kind floor 0.5 → 0.0 | 0 | nothing (diagnostic-only, by design) |

Steps 3 and 6 are the two terms the document leans on hardest in its §3 and §5.3 rows, and
**the suite cannot tell whether they are there.** Both need a unit test (a two-word vs
one-word full match competing for one object; a create and a cancel of the same obligation
in one batch with the object stamped to neither).

Everything else in the suite is sensitive, and the tests assert what the doc says they
assert: `FLOOR`, `MOVE_EMAILS`, `CANCEL_EMAILS` are computed from the corpus, the `real` pin
skips itself on a corpus-hash mismatch (`:236-238`), and the pinned worlds use `==` rather
than `≤`. Two thresholds carry slack (§6.2).

Note the guard suite imports `sb.live.runner`, which imports `httpx` — the offline test
suite now depends on it.

---

## 11. Downstream, and the register's entry condition

| consumer | reads | verdict |
|---|---|---|
| `sb/engine.py:147` | `grade_email(answer, ctx, state, turn).passed` | fine — signature unchanged, delegates to `grade_batch` |
| `sb/live/runner.py:637-641` | `_grade_day` → `results[eid]` | fine — single call site |
| `_print_email` (`runner.py:80-88`) | `d["passed"]/["expected"]/["actual"]/["reason"]` | fine — every detail carries all five keys on all 167 emails |
| capture `verdicts` (`runner.py:625`) | `asdict(EmailResult)` | fine — same dataclass, JSON-serialisable, key set `{passed,max,details,headline}` unchanged |
| `sb/regrade.py:60` | `_grade_day` | fine — 114/167; recomputed `by_eid` matches the recorded one on all 57 days |
| `sb/tests/test_capture_regrade.py:80` | `_grade_day` | fine — live == offline by construction now |
| `sb/analyze.py:25` | `\b(PASS\|FAIL)\b\s+\[(\d+)\]\s+(\S+)` over the log | fine — never reads a reason |
| `sb/schema.py` lint #5 | `op.match` | dead but harmless, as the document says |
| `webapp/scripts/vendor_sb.py:33-34` | copies `grader.py` (not `live/`) | fine — the new grader adds only `import re`; the webapp never calls `grade_email` |
| `sb/oracle.py` | `op.name.replace("_", " ")` | fine — 0 same-node same-kind name-substring collisions; engine 167, scale 167/167 and 287/287 |

**Entry condition** (register `:1991`, *"G-1, G-2, G-7 and the kind filter are one
contract"*) — **met.** All four move in one contract and each is separately evidenced:
G-1 by word-level identity (`Board sign-off` ≡ `board signoff`), G-2 by
`count: too many 11 → 0` with no email flipping to a false pass, G-7 by the Boston case
passing at 0.67 while a genuine survivor still fails, the kind filter by same-kind claiming
plus a reported-not-claimed wrong-kind label. The caveat is §3: G-2's replacement is closed
for same-node stamps only, and the register must say so.

---

## 12. Verdict table

| # | claim | verdict |
|---|---|---|
| 1 | iteration 4 = 114 / 167 / 167 / 137 / 165 / 64 / 64×4 / 67 / 152 / 156 / 77 / 64 / 64 / 166 / 159 | **CONFIRMED** — independent harness, prototype and overlay all agree on every cell |
| 2 | shipped and iteration 3 columns, all 18 | **CONFIRMED**, every cell |
| 3 | op-level profile 76 / 28 / 6 / 21 / 0 / 1 / 2 / 2 / 60 | **CONFIRMED**, every cell |
| 4 | 0 email verdict flips vs iteration 3 | **CONFIRMED** |
| 5 | `oracle_engine` = 167 through `sb.engine.run`; `sb.scale` 167/167 and 287/287 | **CONFIRMED** |
| 6 | invariant under 8 shuffles × 17 worlds | **CONFIRMED and exceeded** — 60 shuffles × 21 worlds × both paths, 0 verdict flips |
| 7 | the staged port is the prototype | **CONFIRMED** — 0 differences in verdict, reason and `actual` over 41 world/path dumps |
| 8 | `_grade_day` is what the runner, `sb.regrade` and the capture test all call | **CONFIRMED** — one call site each; regrade prints 114 |
| 9 | `sb/engine.py` unchanged and consistent with the batch path | **CONFIRMED** on 19 of 21 worlds; the two mis-stamp worlds split, and shipped splits identically |
| 10 | 104 tests, ~2 s | **CONFIRMED** — 104 passed in 1.63 s |
| 11 | rule 3: "a duplicate, whatever id it carries" | **WRONG / UNSAFE AS STATED** — cross-node and unresolvable stamps score **167 with 594 objects**; the three launder worlds all stamp inside the node |
| 12 | "No attribution rule … nothing passes or fails on `email_id`" | **WRONG** — rank step 4 precedes verb priority; a sibling-cancel stamp flips two emails to fail. Step 4 is inert on all 21 worlds |
| 13 | "the verdict cannot depend on the order the store lists objects in" | **WRONG as a property**, CONFIRMED as a measurement — description and `email_id` are absent from the key; two same-titled objects give 2/2 or 0/2 by pool order |
| 14 | batch grading is what closes laundering | **WRONG** — batching is inert (1 op reason, 0 verdicts, 21 worlds); the day-scoped `today` is the whole fix |
| 15 | joint cancel ranking closes `manufacturing-kickoff-3` | **WRONG** — cancel-after is inert everywhere; **title precision** closes it (disable it and `real` = 115), and batching is what opened it |
| 16 | "the cancel matches the walk-through on four words, the move on one" | **WRONG** — title overlap is the first key: the move scores 1.00 there, the cancel 0.25 |
| 17 | `create` claims only today's objects; 8 `wrong day` → `not found` | **OVERSTATED** — the rule is real (and worth 2 points on `wrongkind`), but **7** became `not found` and 1 became `wrong kind` |
| 18 | 6 of 6 wrong-kind labels genuine | **WRONG** — **5 of 6**; `approval-of-budget-tier` reports an earlier day's answer to a different obligation, matched through its description |
| 19 | the 6 include `event day!` | **WRONG** — `event day!` is one of the three that *stopped* being a wrong-kind label |
| 20 | `wrongkind` +3 = right-kind siblings sharing a single content word | **WRONG** — it is three emails with one event op and one todo op of overlapping vocabulary whose objects cross-match when both kinds are swapped |
| 21 | stale floor `> half` catches 11 of 15 retitled copies; the 4 that escape are two-word names | **CONFIRMED**, and robust to dropping the first word instead of the last (156 either way) |
| 22 | the retitle guard's requirement is "computed from the corpus, not asserted" | **OVERSTATED** — computed **10**, actually caught **11**: the formula assumes the dropped word is a keyword and `'Endorsement LOI Review'` ends in a stop word. One point of slack |
| 23 | `launder` requirement `≤ null + 15` | **CONFIRMED as sound**, two points loose (79 vs 77) |
| 24 | the stale rule is free on the oracles now | **CONFIRMED** — `oracle_subject` 137 with and without it (iteration 3 lost 1) |
| 25 | the one stale flag is a description-driven mis-diagnosis in an email that fails anyway | **CONFIRMED** — true-positive rate 0 of 1; and the absolute floor lets a description-only match fail a move whose own claim scored 0.5 |
| 26 | cancel τ = 1.0 is tight; `nocancel` = 167 − 8 | **CONFIRMED** — 0 false-passing cancel ops; the historical G-7 Boston case now passes at 0.67 |
| 27 | cancels are same-kind again, closing iteration 3 §5.4 | **CONFIRMED** |
| 28 | a weaker sibling claim cannot shield a cancel | **CONFIRMED for the iteration-3 shape** (`create_fresh_only` closes it); **NEW EXPOSURE** — exclusive claiming means one survivor can accuse only one cancel: two cancel emails of one node on one day give a false PASS where shipped fails both. Latent (each node has exactly one cancel email) |
| 29 | stemmer fixed; `oracle_inflect` 159 → 165 | **CONFIRMED** — 36 of 207 op-name words still split vs 59 before; `freeze/freezes` is a new split, and `news→new` / `meetings→meet` persist by choice |
| 30 | stop list: 97 words, 66 unused, 7 move a world by ≤ ±2, none moves `oracle_name` | **CONFIRMED** to the digit |
| 31 | `sb/schema.py` rejects a name with no letters or digits | **CONFIRMED**, and the guard's coverage is exactly the empty-keyword case |
| 32 | `sb/oracle.py` can title by `op.name` | **CONFIRMED** — 0 name-substring collisions, engine 167, scale 167/167 and 287/287 |
| 33 | `test_e2e.py:58` does not flip and pins `"stale copy"` | **CONFIRMED** by the passing suite; mutating the stale rule fails it |
| 34 | all 10 downstream consumers unaffected | **CONFIRMED** — plus `_print_email` smoke-tested and the capture `verdicts` shape verified over all 167 emails |
| 35 | store insertion order and duplicate acceptance made iteration 3 §2/§3 reachable live | **CONFIRMED** — `_events` is a dict rendered by `list(_events.values())` (`store_app.py:225-230`), `create_event` has no dedupe (`:127-131`) |
| 36 | the guard set is the adversaries that broke iterations 1 and 3, as tests | **CONFIRMED and materially stronger** — but **`words matched` and verb priority survive mutation with 104/104 green**, and batch grading is detected by exactly one unit test |
| 37 | the register's phase-2 entry condition is met | **CONFIRMED**, with §3's caveat on G-2's scope |
| 38 | there is a minimum standard of evidence for a claim | **not claimed, and worth stating** — a create/move claim has **no overlap floor**; a one-content-word-title agent scores **165** |

---

## 13. What I could not check

1. **Whether any of this is closer to the truth.** Unchanged from the previous two passes.
   Phase 1d is deferred behind C-10, so every number is grader-versus-grader. My audits in
   §7 and §8 are my own reading of the capture, not a human reference.
2. **Generalisation.** One capture, one model, one seed (42), one lever set, one corpus sha
   `03e0d963b9866d8f`. K-2 records 19 of 100 seeds raising `InfeasibleSchedule`, so seed
   variance is still unrunnable, and the document's open issue 6 says the right thing.
3. **Whether a real model would mis-stamp across nodes.** A-5 still records no observed
   instance. §3 is an exploit that exists, not one that has been used — and its likelier
   form is honest transcription error (A-4: 108 of 167 ids exceed 40 chars, 16 prefix
   pairs), which is silently *rewarded* here rather than punished.
4. **Predicate coverage.** Unchanged: 112 `eq`, 12 `by`, 1 `any_of`, 0 `in`/`not_in`,
   `exact_day` on 134/134. Nothing here exercises interval predicates or `within:Nd`.
5. **The live store under concurrency.** I read `store_app.py` but ran no live server, so
   "insertion order" is a code reading plus the shuffle proxy, not an observation.
6. **Whether `open issue 5` (a create and its move in one batch) is really absent in 69 of
   100 seeds.** I confirmed it is absent in *this* plan (0 node-days touch one obligation
   twice); the 69/100 figure I could not reproduce, because 19 of 100 seeds do not build.
7. **The corpus authority question.** C-10 is resolved for working purposes only; the
   contract keys identity on `op.name`, so an adverse answer moves every number here.

---

## 14. Recommendations

Four required, all measured free; then the doc corrections. With these, iteration 4 is
implementable and the register should keep it at `fix proposed` (never `verified` before
phase 1d).

1. **Remove rank step 4 (`mine`), or move it after the verb term.** It changes 0 verdicts
   and 0 scores on all 21 worlds, and it is the only reason a mis-stamp can decide a claim —
   which the document's own "no attribution rule" says must not happen. Add the CASE 1
   construction as a unit test either way.
2. **Make the ranking key actually total: append `(o.description, o.email_id)`.** Free (no
   world contains a colliding pair) and it converts the measured invariance into a property.
   Keep the existing shuffle tests; they will not detect this, so add CASE 4 as a unit test.
3. **Require the wrong-kind report candidate to be in `today`** (or raise the floor to
   0.75). Removes the one false label at zero score cost and takes the label set to 6 of 6
   genuine — which is what the document already claims.
4. **Fix `RETITLE_CAUGHT` by simulating the word drop rather than counting keywords**
   (10 → 11), and tighten the `launder` bound from `null + 15` to the measured 77. Add
   **`launder_cross`** as a world with an explicit requirement — either `≤ null` if the fix
   is taken, or a comment recording that it scores 167 and why.

Then, in `docs/grader-contract.md`, regardless of what ships:

5. **Withdraw "That is a duplicate, whatever id it carries" (`:95`).** Replace with the
   measured scope: same-node stamps only; cross-node and unresolvable stamps leave the pool
   at `runner.py:139-145` and score 167. Say the same in the changelog, and update A-5 in the
   register to record that this contract does **not** close it.
6. **Withdraw "Nothing passes or fails on it" (`:122`)** unless recommendation 1 is taken.
7. **Re-attribute the two fixes.** The laundering fix is the **day-scoped `today`**, not
   batching (batching is inert: 1 op reason, 0 verdicts, 21 worlds). The
   `manufacturing-kickoff-3` closure is **title precision**, not joint cancel ranking
   (cancel-after is inert; without precision `real` = 115). If batching is kept, say plainly
   that it buys nothing measured and costs a shielding exposure held off by a tie-break —
   or drop it and keep the day scope, which is the whole benefit and half the code.
8. **Correct the four counts:** 7 not 8 reclassifications; 5 of 6 wrong-kind labels genuine
   (until fix 3); `event day!` is not among them; and rewrite the `wrongkind` +3 audit — it
   is an opposite-kind two-op swap, not a single-content-word sibling.
9. **State that a create/move claim has no overlap floor**, next to "no `dateok` term". A
   one-word-title agent scores 165 and a half-name agent 157. The +17 is bought partly by
   permissiveness, and a reader who does not know that will over-read it.
10. **Add the two missing guards** (`words matched`, verb priority) — both survive mutation
    with the whole suite green — and, if cheap, a unit test for CASE 5b (two cancels, one
    survivor) and CASE 8 (an object made in a sibling's turn), so the two latent shapes are
    on record before a phase-5 corpus edit makes them live.
