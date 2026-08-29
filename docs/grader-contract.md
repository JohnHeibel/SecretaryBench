# The grader's identity contract — phase 2

**Status: `applied`, iteration 4, revised after its adversarial pass.** Implemented in
`sb/grader.py`, `sb/oracle.py`, `sb/live/runner.py`, `sb/regrade.py`, `sb/schema.py` and the
tests, in one commit with no corpus change. Not `verified`: nothing here can be until the phase
1d hand-grade exists (open issue 9). Three adversarial passes so far:
iteration 1 rejected outright (`docs/_repair/VERIFY-phase2.md`); iteration 3 rejected on two
safety properties that did not hold (`docs/_repair/VERIFY-phase2-iter3.md`); iteration 4
passed with four required changes, all measured free (`docs/_repair/VERIFY-phase2-iter4.md`).
Those four are applied below, together with one simplification the pass offered.

Register: **G-1, G-2, G-3, G-5, G-7, G-8, G-9, G-10, A-5**, and **C-10**.
Evidence: `captures/baseline-sonnet-4-5`. Prototype and guard harness:
`docs/_repair/phase2_guards.py` (`grade_v5`, `run_guards5`, `report5`).

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

## The contract

### 1. Identity from the obligation's name, not a generated keyword

Keywords are the **stemmed content words of `op.name`** (stop-words dropped on the raw word;
falls back to all words when a name is entirely stop-words, e.g. `event day!`). An
(obligation, object) pair scores the fraction of keywords present in `title + description`;
the same fraction over the **title alone** is the primary evidence (G-3: a description can
complete a match, never outrank a title match). There is **no minimum overlap** for a claim:
one shared content word on the right day passes. A perfect scheduler that titles every object
with a single content word of the obligation's name scores 165. Read the +17 with that in mind.

`match` is no longer consulted, so `fix_match.py` becomes unnecessary and the corpus needs no
mutation — which resolves the largest source of the C-10 fork.

### 2. Exclusive assignment, same kind, one ranking

Every op of the email claims at most one **same-kind** object and every object serves at most
one op, by one ranking:

1. title overlap; 2. title + description overlap; 3. created **today**; 4. title precision
(how much of the title *is* the obligation); 5. create/move before cancel; 6. the op's
position; 7. the object's title, description, date and stamp.

A `create` may claim only an object **created today** (it asked for one). A `move` may claim
any. A `cancel` enters the ranking only at **full overlap** — every content word present.

Step 7 makes the key a total order: the verdict cannot depend on the order the store lists
objects in. "Created today" is the day's new-object set, not the `email_id` stamp: the stamp
decides nothing here unless two objects are otherwise identical. (Iteration 4 as first
written ranked "created for this email" ahead of the verb, so a correct object stamped with a
sibling cancel's id failed two emails; `VERIFY-phase2-iter4` §5.3. Removed: 0 verdicts moved.)

### 3. Over-creation

A `create`/`move` **fails** if another same-kind object **created today**, stamped with any
email of this node, matches it at least as well and was claimed by no op of the email. That is a
duplicate. This is the day-scoped `today` the iteration-3 pass recommended; it is what closes
laundering inside a node (`launder_all` / `launder_past` 152 / 139 → **64 / 64**).

What it does **not** close: copies stamped with an email of **another node**, or with an id that
resolves to nothing, never reach this node's pool — `sb/live/runner.py:139-145` drops them by
attribution before the grader sees them. A launderer that does that scores **167 with 594
objects** (`launder_cross`), as it does under the shipped grader at its own ceiling. That is
register **A-5**, which this contract does not close; phase 3 owns attribution. The guard suite
pins the value so the day it moves, someone changed attribution on purpose.

### 4. Stale copies on a move

A `move` **fails** if an unclaimed same-kind object from an **earlier day** still carries more
than half the obligation's words. A move's obligation already had an object; after the move
exactly one may remain (`runner.py:107`, *"Never leave duplicates"*; `test_e2e.py:58`).

### 5. Cancel

A `cancel` **fails** if it claims anything: a same-kind survivor carrying every content word of
the obligation that no other op of the email accounts for. Cross-kind survivors are not its
business (the shipped rule; G-7).

### 6. Diagnosis

A failed `create`/`move` reports `wrong kind: created a <kind>, expected a <kind>` when an
unclaimed object of the other kind carries more than half its words, some in the title, and
(for a create) was made today; else `no <kind> titled like "…" was created`. Other reasons:
`over-created: N equally-matching …`, `moved, but N stale copy left behind (double-booked)`,
`should be cancelled, but still on the calendar`, `on the wrong day`.

### Explicitly not included

- **No `dateok` term anywhere.** Grading must not search the model's output for something that
  fits (iteration 1; 63 points of gaming headroom).
- **No attribution rule in the grader.** `email_id` decides only which email's *turn* an object
  belongs to — the no-action check — and, upstream, which node's pool it enters (A-5, above).
- **No batch assignment.** Iteration 4 first graded a node's emails for one day together. Its
  pass measured that inert (1 op reason string on 21 worlds, 0 verdicts) and found it had
  *created* a cross-email exposure — a sibling's move could take a cancel's survivor — that a
  tie-break was holding off. Dropped; the day-scoped `today` is the whole benefit.
- **No `words matched` term.** For a title-overlap tie on one object it orders exactly as title
  precision does (both are `|keywords| / |title words|` there); mutation-testing found nothing
  detected its removal. Dropped.
- **No kind tolerance.** K-7's four mis-keyed ops are a corpus edit for phase 5.
- **No optimal assignment.** Greedy loses 2 create/move ops on the capture to optimal; optimal
  fabricates a pass (`VERIFY-phase2` §3).

---

## What the adversarial passes found, in order

| pass | killed | answered by |
|---|---|---|
| 1 (`VERIFY-phase2`) | no volume brake: a date-blind shotgun scored 148, five copies of everything 167; `dateok` tie-break searched the output; oracle circular | rules 3–4, the guard set, stemming |
| 2 (`VERIFY-phase2-iter3`) | the brake was scoped to the `email_id` split on the runner path (a launderer scored 167 with 509 objects; the engine path said 64); assignment order-dependent (`oracle_name` 167 in 19 of 60 orderings); cross-kind claiming made cancel harder and shielded duplicates; stale rule evaded by one retitled word | day-scoped `today`; total-order key; same-kind claiming; stale floor `> half`; stemmer |
| 3 (`VERIFY-phase2-iter4`) | "a duplicate whatever id it carries" false across nodes; `mine` ranked before verb so a stamp decided verdicts; key not total (description absent); batch grading and joint cancel ranking inert and misattributed; four audit counts wrong; retitle threshold 10 for 11 caught; `words matched` and verb priority unguarded | `mine` → "created today"; key tail `(title, description, when, stamp)`; wrong-kind report needs a today object for creates; thresholds computed by simulation; batching dropped; the counts corrected below; unit tests for verb priority, specificity, determinism-with-descriptions, the sibling-cancel stamp |

---

## The guard set

Built from the adversaries that broke iterations 1 and 3. All free, offline, seconds. Every
world is scored on **both** grading paths (day-end state split by `email_id`, and per-email
snapshots as `sb.engine` sees them) and under **pool shuffles**.

| guard | what it is | requirement |
|---|---|---|
| `oracle_engine` | `sb/oracle.py` titled by `op.name`, through `sb.engine.run` | **167** — `sb/scale.py:127` gates on it |
| `oracle_name` | perfect agent, titles = obligation name | **167** |
| `oracle_subject` / `oracle_inflect` | perfect agent, subject-line titles / pluralised titles | pinned |
| `null` | creates nothing | **64** (register V-3) |
| `dup5` · `shot7/45/90` | five copies of everything · date-blind, one object per day over N days | **≤ null** |
| `dupmove` | every move leaves the old object behind | **≤ oracle_name − 15** (the move emails) |
| `dupmove_retitle` | as above, one word dropped from each stale copy's title | **≤ oracle_name − 11** (copies keeping more than half the words; simulated, not counted) |
| `launder` | right answer + 5 copies of every *created* object, stamped with a node sibling's id | **≤ null + 13** (its move-only emails) |
| `launder_all` · `launder_past` | copies on moves too · copies stamped with an already-graded sibling | **≤ null** |
| `launder_cross` | copies stamped with **another node's** email id | informational: **167**, A-5, pinned |
| `nocancel` | perfect, never deletes anything | **167 − 8** (the all-cancel emails) |
| `wrongkind` · `bothkinds` | right title and day, wrong kind · an event *and* a to-do for every obligation | pinned |
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
| `launder_cross` | 158 | — | 167 (A-5, see rule 3) |
| `bothkinds` | 153 | 166 | 166 |
| `nocancel` | 150 ✗ | 159 | **159** |
| engine == runner on every world | yes | **no** (launder: 77/167, 64/152, 64/139) | **yes** |
| invariant under pool shuffles | yes | **no** (`oracle_name` 165–167, `dupmove` 151–152, …) | **yes** (60 × 21 worlds × 2 paths, third pass) |
| **guards** | fails 5 | fails 4 + 2 properties | **PASS** |

The third pass reproduced every cell on an independent implementation and found the staged
port identical to the prototype at verdict, reason and `actual` level on 21 worlds × 2 paths.

**Op-level failure profile on the real capture** (190 ops):

| bucket | shipped | iter. 3 | **iter. 4** |
|---|---|---|---|
| ok | 56 | 76 | **76** |
| not found | 50 | 18 | 28 |
| wrong kind | — | 8 (2 false) | **6** (0 false after the today rule; was 5 of 6) |
| wrong day | 14 | 29 | 21 |
| count: too many | 11 | 0 | 0 |
| stale copy on move | — | 1 | 1 |
| cancel residue | 3 | 2 | 2 |
| over-acted | 2 | 2 | 2 |
| **failing ops** | **80** | **60** | **60** |

Iteration 4 changes **no email verdict** against iteration 3 (0 flips). Seven `wrong day`
diagnoses became `not found` and one became `wrong kind`: they were creates claiming an
inherited sibling's object (`'Design Lead 1:1'` → the design walk-through, `'Actnano Visit'` → a
strategy meeting). The model created nothing for those obligations that day, and the grader
now says so.

**Read the `wrong day` count honestly.** Roughly half the work the identity rule recovers turns
out to be on the wrong date — error that was always there and invisible behind "couldn't find
it". The +17 is net of it.

### Audits on the real capture

- **Both remaining cancel failures are genuine** (confirmed by all three passes). In
  `manufacturing-kickoff-3` the survivor is the model's own answer to the very obligation being
  cancelled, matched through its description; the shipped grader passed it falsely. Under
  batch grading a sibling's single-word `move 'Reveal Event'` briefly took that survivor and
  the cancel passed; **title precision** is what steered the move to `"WC cleat reveal event"`
  (0.33 vs 0.20), and without batching the exposure does not exist.
- **The one stale flag is a mis-diagnosis inside an email that fails anyway**:
  `reveal-event-date-and-venue-3`, `move 'Reveal Event'` — `event` is a stop word, so the keyword
  set is `{reveal}` and two siblings (the press briefing by description, the walk-through by
  title) read as copies. The stale rule's true-positive rate on this capture is 0 of 2 flags;
  its justification is the `dupmove` guards, on which it is exact.
- **The 6 `wrong kind` labels** are genuine model-vs-key disagreements, K-7's `Team pizza
  party` (keyed `todo`) among them. `event day!` is not among them: its `"Event Pop Up"` was a
  loose cross-kind claim in iteration 3 and is now `not found`.
- **`wrongkind` +3 over the floor**: three emails with one event op and one to-do op of
  overlapping vocabulary; when every kind is swapped, each op's object satisfies the other's
  on the right day. Kind-filtered claiming shows the same.
- **Stop-word sensitivity**: leave-one-out over all 97 words moves some world for 7 words, never
  by more than ±2, none moves `oracle_name`. 66 of the 97 never occur in an op name.

### Blocker closures of iteration 2, as they stand now

| blocker | resolution |
|---|---|
| kind filter | same-kind claiming (the shipped semantics, exactly); wrong-kind reported with a floor; K-7 surfaces as `wrong kind` |
| `cancel` bypasses the assignment | in the assignment at full overlap; a sibling op's object is claimed by the sibling, and a perfect tie goes to the work |
| `sb/oracle.py:52` titles by `match` | becomes `op.name.replace('_', ' ')`, same commit as the grader; 167 through the engine and through `sb.scale` (`--filler 0` and default) |
| `test_e2e.py:58` would flip | it does not: the double-booked reschedule fails under rule 4; the test pins the reason |

---

## Open issues

1. **Single-content-word obligations** (24 of 134) remain the soft spot of every rule: the
   only defence on them is the date predicate, the stale rule mis-describes siblings that share
   the word, and one shared word on the right day is a pass. The G-5 **name-aware lint** (flag
   same-kind siblings whose keyword sets nest, and one-word names) belongs to phase 5 with the
   renames it will demand. `sb/schema.py` lint #5 still checks the old `match` rule; dead once
   `match` is unread, and replacing it fails the current corpus on 10 names.
2. **Cross-node and unresolvable stamps** are invisible to every rule (A-5; rule 3). Phase 3.
3. **A same-day sibling whose object contains every keyword of this op** reads as a duplicate
   under rule 3, because each email is graded alone against the day's objects: `create 'Board
   Sync'` (`sync` is a stop word, so `{board}`) next to a sibling's correct `"Board memo
   review"` fails as over-created. Batch grading would have let the sibling claim its own
   object; it was dropped for the cross-email cancel exposure it created. Both need G-5
   nesting, and no feasible seed of 100 batches a nested pair on one day; the lint in issue 1
   removes the condition.
4. **Two-word names** escape the retitled-stale-copy check by construction (4 of 15 move
   emails). Any lower floor breaks the oracles (`≥ half`: 166 / 166).
5. **Two cancel emails of one node on one day** would let one survivor accuse only one of them
   (exclusive claiming); the other passes falsely. Every node has exactly one cancel email
   today; a phase-5 edit could change that.
6. **`event day!`** is the only all-stop-word name; its fallback keywords are words every other
   op stops. Rename in phase 5.
7. **Descriptions** stay in identity (title first). The `manufacturing-kickoff-3` cancel is
   right *because* of one; the `Reveal Event` stale flag is wrong because of one.
8. **One capture, one model, one seed, one lever set.** The guards are model-independent; the
   +17 is not. It needs a second captured run before anyone quotes it.
9. **Everything here is grader-versus-grader.** Only the phase 1d hand-grade can say whether the
   score moved toward *truth*, and it is deferred behind C-10. Per the register's status legend
   this can reach `fix proposed` — never `verified`.

---

## Implementation notes

- `grade_email(answer, ctx, state, turn, today=None)`: the fifth argument is everything created
  today; it defaults to `turn`, which is exact on the engine path. `sb/engine.py` is unchanged.
- `sb/live/runner.py` gains `_grade_day(corpus, plan, batch, state, day_new)`: splits the day's
  new objects by stamp into each email's `turn`, passes the whole day as `today`. The runner's
  day loop, `sb/regrade.py` and `test_capture_regrade.simulate` all route through it.
- `EmailResult`/`details` keep their shape; all ten downstream consumers checked by the passes.
- `sb/oracle.py` titles by `op.name.replace('_', ' ')`. `sb/schema.py` rejects a name with no
  letters or digits and its `Op` docstring stops calling `match` the grader's field.
- `sb/tests/test_e2e.py:58` keeps `assert not passed` and pins `"stale copy"`.
- `sb/tests/test_grader_guards.py` is the guard set above as tests — 18 worlds, both paths,
  shuffles, thresholds computed from the corpus by simulation, and pins for `real` (114),
  `oracle_subject` (137), `oracle_inflect` (165), `wrongkind` (67), `bothkinds` (166),
  `launder_cross` (167). The `real` pin skips itself if the corpus no longer matches the
  capture's hash. `sb/tests/test_grader.py` gains 16 identity tests (G-10), including the two
  the third pass found unguarded by mutation (verb priority, specificity). Whole suite ~2 s.
- `webapp/api/_lib/sb/` is regenerated at build by `webapp/scripts/vendor_sb.py`; nothing to do.
