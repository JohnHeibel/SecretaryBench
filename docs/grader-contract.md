# The grader's identity contract — phase 2

**Status: `fix proposed`, iteration 4.** Designed and measured, **not implemented**.
`sb/grader.py` is unchanged. Iteration 1 was rejected outright by an adversarial pass
(`docs/_repair/VERIFY-phase2.md`); iteration 2 left four blockers; iteration 3 closed them and
was rejected by a second pass (`docs/_repair/VERIFY-phase2-iter3.md`) on two load-bearing
safety properties that did not hold. This iteration answers that pass, item by item, and is
awaiting its own before anything is written into `sb/`.

Register: **G-1, G-2, G-3, G-5, G-7, G-8, G-9, G-10, A-5**, and **C-10**.
Evidence: `captures/baseline-sonnet-4-5`. Prototype and guard harness:
`docs/_repair/phase2_guards.py` (`grade_batch_v5` / `grade_v5`, `run_guards5`, `report5`).

---

## The problem in one line

The grader cannot recognise a correct answer, so scores measure naming luck rather than
scheduling skill.

Three defects stack:

| # | defect | register |
|---|---|---|
| 1 | Identity is a **literal substring** match against a keyword the model is never told. `"Board sign-off meeting"` fails `"signoff"` on one hyphen. | G-1 |
| 2 | Those keywords **are not authored** — `scripts/fix_match.py` generated them from obligation names across 96 emails. They do not exist upstream. | C-10 |
| 3 | The **exactly-one rule** fires on *distinct* obligations that share vocabulary, not on duplicates. 0 of 57 recorded "duplicates" were duplicates. | G-2 |

Defect 1 also hides defect 4: the grader usually never reached the date check, so date errors
were invisible behind "couldn't find it".

**Proof it is the grader, not the models:** a scheduling-perfect agent that titles each object
by the obligation's own name scores **158/167** under today's grader. All nine losses are the
grader's — two whole-name substrings (G-8), seven exactly-one collisions between distinct
obligations (G-2). The benchmark cannot award full marks to a flawless assistant.

---

## What the second adversarial pass found, and what changed

`docs/_repair/VERIFY-phase2-iter3.md` confirmed every number of iteration 3 to the digit and
then broke two of its stated safety properties:

| finding | iteration 3 | iteration 4 |
|---|---|---|
| **§2** The volume brake was scoped to "the turn", but on the runner and regrade paths the turn *is* the `email_id` split. Duplicates stamped with a sibling's id were invisible to it: a launderer scored **167** with 509 objects in the store, and the same state scored 64 on the engine path. | ✗ | The emails of a node served on one day are graded **together** (`grade_batch`) against the day's new objects. A sibling's correct object is claimed by the sibling; what nobody claims is surplus. Laundering scores **64 / 64 / 77** on three variants; both paths agree on every world |
| **§3** The greedy assignment had no tie-break for nested keyword sets, so `oracle_name` = 167 held in 19 of 60 pool orderings and a perfect agent that handled two moves in the other order read 165 through `sb.engine.run`. | ✗ | The sort key is a **total order on content** (title overlap, full overlap, words matched, created-for-this-email, title precision, verb, then title and date). Invariant under 8 shuffles on all 17 worlds |
| §5.3 A `create` could claim an inherited object, so a same-email cancel of it passed with the model doing nothing (constructed). Under batch grading the same shape appeared **on the real capture**: a sibling's single-word `move 'Reveal Event'` took `"Design walk-through at WC reveal"` and shielded the cancel in `manufacturing-kickoff-3`. | latent | A create claims only an object **made today**; cancels are ranked **jointly** with create/move so the op that matches more of an object's words gets it (the cancel matches the walk-through on four words, the move on one). The false pass is closed on both shapes |
| §5.4 / §6 The cross-kind pool made `cancel` strictly harder than the shipped rule, let a wrong-kind claim shield a duplicate, and the stated reason for rejecting overlap-first ordering did not reproduce. | ✗ | Claiming is **same-kind** again, for every verb. A wrong-kind object is **reported, never claimed**, and only when it carries more than half the obligation's words in title + description, some of them in its title, and no sibling op claimed it. 6 of 6 such labels on the capture are genuine (were 6 of 8) |
| §4.1 The stale rule (`≥ claimed score`) was evaded by dropping one word from the stale copy's title (163 vs a 152 guard). | ✗ | Floor is **more than half the obligation's words**: 11 of 15 retitled copies caught (was 4). The 4 that escape are two-word names losing exactly half, by construction; the guard's requirement is computed from the corpus, not asserted |
| §7 Stemmer over-stripped `-es` and broke `note/notes`, `trophy/trophies`; stop list applied pre-stem. | | `-ies → y`, `-es` only after a sibilant, `-s` not after `s`: `oracle_inflect` 159 → **165**. Stopping on the stem too was measured: −1 on `oracle_subject`, 0 elsewhere, and it collapses `news` → `new` → stopped; **not applied**, stop-words stay matched on the raw word |
| §7 A name with no letters or digits has an empty keyword set: its cancel passes vacuously. | latent | `sb/schema.py` raises `CorpusError` |
| §11.5 / §11.7 (could not check) | | `store_app.py` keeps insertion order and delete-then-recreate moves an object to the end, so §3 was reachable live; the store accepts exact-duplicate creates, so §2 was reachable live. Both are now guarded, not assumed |

Also corrected in this document: the shipped grader's `dupmove` 148 is a guard **failure**
(it misses 5 of 15 double-bookings), and the claim that turn membership "cannot be defeated
by a mis-stamped `email_id`" is withdrawn — the mechanism that makes A-5 harmless to the
brake is batch grading, and it is guarded by three launder worlds.

---

## The contract

### 1. Identity from the obligation's name, not a generated keyword

Keywords are the **stemmed content words of `op.name`** (stop-words dropped on the raw word;
falls back to all words when a name is entirely stop-words, e.g. `event day!`). An
(obligation, object) pair scores the fraction of keywords present in `title + description`;
the same fraction over the **title alone** is the primary evidence (G-3: the description can
complete a match, never outrank a title match).

`match` is no longer consulted, so `fix_match.py` becomes unnecessary and the corpus needs no
mutation — which resolves the largest source of the C-10 fork.

### 2. One assignment for the day, per node

The emails of a node served on one day are graded together. Every `create`/`move`/`cancel` op
of those emails claims at most one **same-kind** object and every object serves at most one
op, by one ranking:

1. title overlap; 2. title + description overlap; 3. words matched (a two-word full match
beats a one-word full match); 4. created for this email; 5. title precision (how much of the
title *is* the obligation); 6. create/move before cancel; 7. the op's position; 8. the object's
title and date.

A `create` may claim only an object **created today** (it asked for one). A `move` may claim
any. A `cancel` enters the ranking only at **full overlap** — every content word present.

Steps 7–8 make the ranking a total order on content: the verdict cannot depend on the order the
store lists objects in, on either grading path. On the engine path each email is its own batch.

### 3. Over-creation

A `create`/`move` **fails** if another same-kind object **created today** matches it at least as
well and was claimed by no op in the batch. That is a duplicate, whatever id it carries.

### 4. Stale copies on a move

A `move` **fails** if an unclaimed same-kind object from an **earlier day** still carries more
than half the obligation's words. A move's obligation already had an object; after the move
exactly one may remain (`runner.py:107`, *"Never leave duplicates"*; `test_e2e.py:58`).

### 5. Cancel

A `cancel` **fails** if it claims anything: a same-kind survivor carrying every content word of
the obligation that no sibling op accounts for. Cross-kind survivors are not its business (the
shipped rule; G-7).

### 6. Diagnosis

A failed `create`/`move` reports `wrong kind: created a <kind>, expected a <kind>` when an
unclaimed object of the other kind carries more than half its words, some in the title; else
`no <kind> titled like "…" was created`. Other reasons: `over-created: N equally-matching …`,
`moved, but N stale copy left behind (double-booked)`, `should be cancelled, but still on the
calendar`, `on the wrong day`.

### Explicitly not included

- **No `dateok` term anywhere.** Grading must not search the model's output for something that
  fits (iteration 1; 63 points of gaming headroom).
- **No attribution rule.** `email_id` decides only which email's *turn* an object belongs to
  (the no-action check and rank step 4). Nothing passes or fails on it.
- **No kind tolerance.** K-7's four mis-keyed ops are a corpus edit for phase 5.
- **No optimal assignment.** Greedy loses 2 create/move ops on the capture to optimal; optimal
  fabricates a pass (`VERIFY-phase2` §3).

---

## The guard set

Built from the adversaries that broke iterations 1 and 3. All free, offline, seconds.
Every world is scored on **both** grading paths (day-end state split by `email_id`, and
per-email snapshots as `sb.engine` sees them) and under **pool shuffles**.

| guard | what it is | requirement |
|---|---|---|
| `oracle_engine` | `sb/oracle.py` titled by `op.name`, through `sb.engine.run` | **167** — `sb/scale.py:127` gates on it |
| `oracle_name` | perfect agent, titles = obligation name | **167** |
| `oracle_subject` / `oracle_inflect` | perfect agent, subject-line titles / pluralised titles | pinned |
| `null` | creates nothing | **64** (register V-3) |
| `dup5` · `shot7/45/90` | five copies of everything · date-blind, one object per day over N days | **≤ null** |
| `dupmove` | every move leaves the old object behind | **≤ oracle_name − 15** (the move emails) |
| **`dupmove_retitle`** | as above, one word dropped from each stale copy's title | **≤ oracle_name − 10** (copies keeping more than half the words; computed) |
| **`launder`** | right answer + 5 copies of every *created* object, stamped with a node sibling's id | **≤ null + 15** (its moves leave nothing) |
| **`launder_all`** · **`launder_past`** | copies on moves too · copies stamped with an already-graded sibling | **≤ null** |
| **`nocancel`** | perfect, never deletes anything | **167 − 8** (the all-cancel emails) |
| `wrongkind` · **`bothkinds`** | right title and day, wrong kind · an event *and* a to-do for every obligation | pinned |
| every world | engine path == runner path; invariant under shuffles | **required** |

The synthetic agents act by obligation identity (they remember which record they made for
each obligation), as a real agent does with the ids the store returns.

---

## Measured

`sb/grader.py` unchanged; contracts applied via the harness over the certified capture. Every
oracle row runs `sb/oracle.py` titled by `op.name` (shipped with its `match` titles: 167).

| world | shipped | iter. 3 | **iter. 4** |
|---|---|---|---|
| real (certified `claude-sonnet-4-5`) | 97 | 114 | **114** |
| `oracle_engine` | 158 ✗ | 167 | **167** |
| `oracle_name` | 158 ✗ | 167 | **167** |
| `oracle_subject` | 95 | 137 | **137** |
| `oracle_inflect` | 142 | 159 | **165** |
| `null` | 64 | 64 | **64** |
| `dup5` · `shot7` · `shot45` · `shot90` | 64 | 64 | **64** |
| `wrongkind` | 65 | 69 | 67 |
| `dupmove` | 148 ✗ | 152 | **152** |
| `dupmove_retitle` | 146 | 163 ✗ | **156** |
| `launder` | 72 | **167** ✗ | **77** |
| `launder_all` | 64 | 152 ✗ | **64** |
| `launder_past` | 64 | 139 ✗ | **64** |
| `bothkinds` | 153 | 166 | 166 |
| `nocancel` | 150 ✗ | 159 | **159** |
| engine == runner on every world | yes | **no** (launder: 77/167, 64/152, 64/139) | **yes** |
| invariant under pool shuffles | yes | **no** (`oracle_name` 165–167, `dupmove` 151–152, …) | **yes** (8 shuffles × 17 worlds) |
| **guards** | fails 5 | fails 4 + 2 properties | **PASS** |

**Op-level failure profile on the real capture** (190 ops):

| bucket | shipped | iter. 3 | **iter. 4** |
|---|---|---|---|
| ok | 56 | 76 | **76** |
| not found | 50 | 18 | 28 |
| wrong kind | — | 8 (2 false) | **6** (0 false) |
| wrong day | 14 | 29 | 21 |
| count: too many | 11 | 0 | 0 |
| stale copy on move | — | 1 | 1 |
| cancel residue | 3 | 2 | 2 |
| over-acted | 2 | 2 | 2 |
| **failing ops** | **80** | **60** | **60** |

Iteration 4 changes **no email verdict** against iteration 3 (0 flips). Eight `wrong day`
diagnoses became `not found`: they were creates claiming an inherited sibling's object
(`'Design Lead 1:1'` → the design walk-through, `'Actnano Visit'` → a strategy meeting). The
model created nothing for those obligations that day, and the grader now says so.

**Read the `wrong day` count honestly.** Roughly half the work the identity rule recovers turns
out to be on the wrong date — error that was always there and invisible behind "couldn't find
it". The +17 is net of it.

### Audits on the real capture

- **Both remaining cancel failures are genuine** (confirmed by both verifiers). In
  `manufacturing-kickoff-3` the survivor is the model's own answer to the very obligation being
  cancelled, matched through its description; the shipped grader passed it falsely.
- **The one stale flag is a mis-diagnosis inside an email that fails anyway**:
  `reveal-event-date-and-venue-3`, `move 'Reveal Event'` — `event` is a stop word, so the keyword
  set is `{reveal}` and the press briefing (description "week before reveal") reads as a copy.
  The stale rule's true-positive rate on this capture is 0 of 1; its justification is the
  `dupmove` guards, on which it is exact.
- **The 6 `wrong kind` labels** are all genuine model-vs-key disagreements, including K-7's
  `Team pizza party` (keyed `todo`) and the all-stop-word `event day!`.
- **`wrongkind` +3 over the floor**: three emails where a right-kind sibling object shares the
  obligation's single content word and sits on the right day. Kind-filtered claiming shows the
  same; it is the single-word identity weakness, not a kind rule.
- **Stop-word sensitivity**: leave-one-out over all 97 words moves some world for 7 words, never
  by more than ±2, none moves `oracle_name`. 66 of the 97 never occur in an op name.

### Blocker closures of iteration 2, as they stand now

| blocker | resolution |
|---|---|
| kind filter | same-kind claiming (the shipped semantics, exactly); wrong-kind reported with a floor; K-7 surfaces as `wrong kind` |
| `cancel` bypasses the assignment | ranked jointly with create/move at full overlap; a sibling's object is claimed by the sibling; a weaker sibling claim cannot shield a cancel |
| `sb/oracle.py:52` titles by `match` | becomes `op.name.replace('_', ' ')`, same commit as the grader; 167 through the engine and through `sb.scale` (`--filler 0` and default) |
| `test_e2e.py:58` would flip | it does not: the double-booked reschedule fails under rule 4; the test pins the reason |

---

## Open issues

1. **Single-content-word obligations** (24 of 134) remain the soft spot of every rule: the
   only defence on them is the date predicate, the stale rule mis-describes siblings that share
   the word, and the `wrongkind` guard sits 3 above the floor because of them. The G-5
   **name-aware lint** (flag same-kind siblings whose keyword sets nest, and one-word names)
   belongs to phase 5 with the renames it will demand. `sb/schema.py` lint #5 still checks the
   old `match` rule; dead once `match` is unread, and replacing it fails the current corpus on
   10 names, so it cannot ride in the grader commit.
2. **Two-word names** escape the retitled-stale-copy check by construction (5 of 15 move
   emails). Any lower floor breaks the oracles (`≥ half`: 166 / 166).
3. **`event day!`** is the only all-stop-word name; its fallback keywords are words every other
   op stops. Rename in phase 5.
4. **Descriptions** stay in identity (title first). The `manufacturing-kickoff-3` cancel is
   right *because* of one; the `Reveal Event` stale flag is wrong because of one. Net on the
   capture: blanking every description moves the score by one email either way.
5. **Same-day ops on one obligation** (a create and its move in one batch) would compete for
   one object. In 69 feasible seeds of 100 it never happens; a corpus edit could make it happen.
6. **One capture, one model, one seed, one lever set.** The guards are model-independent; the
   +17 is not. It needs a second captured run before anyone quotes it.
7. **Everything here is grader-versus-grader.** Only the phase 1d hand-grade can say whether the
   score moved toward *truth*, and it is deferred behind C-10. Per the register's status legend
   this can reach `fix proposed` — never `verified`.

---

## Implementation notes

- `grade_email(answer, ctx, state, turn)` keeps its signature and is `grade_batch([one], turn)`.
  New: `grade_batch(items, today)`, where `items` are `(answer, ctx, state, turn)` for the emails
  of one node served on one day, and `today` is everything created that day.
- `sb/live/runner.py` gains `_grade_day(corpus, plan, batch, state, day_new)`, which splits the
  day's new objects by stamp (each email's `turn`), groups the batch by node and calls
  `grade_batch`. The runner's day loop, `sb/regrade.py` and `test_capture_regrade.simulate` all
  route through it, so the three cannot drift. `sb/engine.py` is unchanged.
- `EmailResult`/`details` keep their shape; `_print_email`, the capture's `verdicts`,
  `sb.regrade` and `sb.analyze` read them unchanged (checked by the verifier, all 10 consumers).
- `sb/oracle.py` titles by `op.name.replace('_', ' ')`. `sb/schema.py` rejects a name with no
  letters or digits and its `Op` docstring stops calling `match` the grader's field.
- `sb/tests/test_e2e.py:58` keeps `assert not passed` and pins `"stale copy"`.
- `sb/tests/test_grader_guards.py` is the guard set above as tests — 17 worlds, both paths,
  shuffles, and pins for `real` (114), `oracle_subject` (137), `oracle_inflect` (165),
  `wrongkind` (67), `bothkinds` (166) — so no future change can quietly move the score. The
  `real` pin skips itself if the corpus no longer matches the capture's hash.
  `sb/tests/test_grader.py` gains 13 identity tests (G-10). Whole suite: 104 tests, ~2 s.
- `webapp/api/_lib/sb/` is regenerated at build by `webapp/scripts/vendor_sb.py`; nothing to do.
