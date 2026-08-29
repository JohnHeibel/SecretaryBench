# The grader's identity contract — phase 2

**Status: `fix proposed`, iteration 3.** Designed and measured, **not implemented**.
`sb/grader.py` is unchanged. Iteration 1 was rejected outright by an adversarial pass
(`docs/_repair/VERIFY-phase2.md`); iteration 2 fixed what it found and left four blockers open;
this iteration closes them and is awaiting its own adversarial pass before anything is written
into `sb/`.

Register: **G-1, G-2, G-5, G-7, G-8, G-9**, and **C-10**. Evidence: `captures/baseline-sonnet-4-5`.
Prototype and guard harness: `docs/_repair/phase2_guards.py` (`grade_v4`).

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
by the obligation's own name scores **158/167** under today's grader. The benchmark cannot
award full marks to a flawless assistant.

> *Corrected in iteration 3.* Earlier drafts said 148. Ten of those nineteen losses were the
> synthetic agent's own fault — it created a new object on every `move` and never removed the
> old one, so the shipped count rule was right to fail it. The agent now moves and cancels by
> obligation identity, as a real agent does with the ids the store returns. The claim stands at
> 158; the nine remaining losses are the substring rule's.

---

## The contract

Five rules. They must land together — measured separately, each looks like it does nothing,
because fixing one relocates failures into another.

### 1. Identity from the obligation's name, not a generated keyword

Keywords are the **stemmed content words of `op.name`** (stop-words dropped; falls back to all
words when a name is entirely stop-words, e.g. `event day!`). Matching is word-level against
`title + description`; the score of an (obligation, object) pair is the fraction of the
obligation's keywords present.

`match` is no longer consulted, so `fix_match.py` becomes unnecessary and the corpus needs no
mutation — which resolves the largest source of the C-10 fork. Stemming is required: without
it word matching is *less* inflection-robust than the substring rule it replaces.

### 2. Exclusive assignment over a cross-kind pool

Every `create`/`move` op scores every object in the node's pool — **both kinds**. Pairs are
sorted and greedily assigned so each object serves at most one obligation and each obligation
claims at most one object. Tie-breaks, in order:

1. **kind matches** — a right-kind object always beats a wrong-kind one, whatever the overlap;
2. **overlap score**;
3. **created this turn** — membership in the turn delta, tested by value
   `(kind, title, when, email_id)`.

*Why kind first.* This makes the assignment score-identical to the kind-filtered pool the
shipped grader uses (measured: identical on every guard column), so the kind rule the model was
given — *"create_event … the calendar; create_todo … tasks that have a deadline"* — is
preserved exactly. What changes is that a wrong-kind object is now **claimed when nothing of the
right kind matches**, and fails with its own reason (`wrong kind: created a to-do, expected an
event`) instead of reading as "nothing created". That closes the kind half of G-9 without
making anything harder. Overlap-first ordering was measured and rejected: it fails a model that
creates both an event and a to-do for one obligation, which the shipped grader passes.

*Why "created this turn" and not `email_id`.* Iteration 2 tie-broke on `email_id == eid`. The
engine path (`sb/engine.py:147`, which `sb.scale`'s gate and `test_e2e.py` run through) passes
no email id, so the tie-break was dead there and **the two grading paths disagreed on the same
state**: `test_e2e`'s double-booked reschedule failed via the engine and passed via the runner-style
per-email split (measured both ways), the engine verdict being an accident of pool order. Turn membership exists on every path and cannot be defeated
by a mis-stamped `email_id` (register A-5).

### 3. A volume brake, scoped to the turn

After assignment, a `create`/`move` **fails** if another **same-kind** object created **in the
same turn** matches it at least as well and was claimed by nobody. That is a duplicate.
Unchanged from iteration 2, where it took `dup5` from 167 to 64 and the date-blind shotguns
from 93–148 to 64.

### 4. A stale-survivor rule for `move`

A `move` additionally **fails** if an **inherited** (earlier-turn) same-kind object matches it
at least as well and was claimed by nobody. A move's obligation already had an object; after the
move exactly one may remain. The prompt says so in as many words (`runner.py:107`, *"Never
leave duplicates"*), `sb/tests/test_e2e.py:58` asserts it, and without this rule a model that
never updates or deletes anything scores a perfect 167 (`dupmove` guard). Create ops are not
subject to it: an inherited object matching a fresh create is a sibling obligation's answer,
which is the whole of G-2.

Both rules judge over-creation on the same haystack as identity. A title-only haystack was
measured: it changes no verdict on any world and makes the threshold loose whenever the claimed
object matched through its description.

### 5. `cancel` goes through the assignment

Create/move ops claim first. Then each `cancel` claims the best **unclaimed** object, of either
kind, whose overlap is **1.0** — every content word of the obligation present. A claim fails the
cancel; no claim passes it. A sibling op in the same email keeps the object it was supposed to
keep, and a stray shared word no longer counts (iteration 1's any-word rule made G-7 worse).

The threshold was swept. Any-word and majority (0.5) both produce false negatives in the oracle
worlds (`oracle_engine` 164 / 166) and cost two real-capture emails each. At 1.0, both remaining
cancel failures on the real capture are genuine — see the audit below.

### Explicitly not included

- **No `dateok` term in the tie-break.** Grading must not search the model's output for
  something that fits. Removing it cost nothing and closed 63 points of gaming headroom (iter. 2).
- **No attribution filter.** `email_id` is not consulted at all now. A-5 stays where it is.
- **No kind tolerance.** K-7's four mis-keyed ops are a corpus edit for phase 5, not a grader rule.

---

## The guard set

Built from the adversaries that broke iteration 1, plus one that broke the synthetic agent
in iteration 3. All free, all offline, all against the 1c capture's corpus and plan.

| guard | what it is | requirement |
|---|---|---|
| `oracle_engine` | `sb/oracle.py` titled by `op.name`, through `sb.engine.run` | **167** — `sb/scale.py:127` gates on it |
| `oracle_name` | perfect agent, titles = obligation name | **167** |
| `oracle_subject` | perfect agent, titles = email subject | high — realistic naming |
| `oracle_inflect` | perfect agent, pluralised titles | high — inflection robustness |
| `null` | creates nothing at all | **64** exactly (register V-3) |
| `dup5` | right answer, then 5 copies of every object | **must not exceed null** |
| `shot7/45/90` | never reads a date; one object per day over N days | **must not exceed null** |
| `wrongkind` | right title and date, wrong kind | informational |
| **`dupmove`** | right answer, but every `move` creates a copy and leaves the old object | **must not exceed `oracle_name` − 15** (the 15 emails with a move op) |

The synthetic agents move and cancel by obligation identity (they remember which record they
made for each obligation). The iteration-2 agent deleted by keyword overlap and never removed
anything on a move, so it *was* the `dupmove` adversary without anyone noticing.

`null` and a uniform date shift are structurally incapable of separating two contracts on this
corpus (56 no-action + 8 all-cancel emails; `exact_day` on 134 of 134 ops). They are retained as
regression tripwires, not as evidence.

---

## Measured

`sb/grader.py` unchanged; contracts applied via the harness over the certified capture.
`shipped` runs `sb/oracle.py` as it is (titled by `match`); every other row runs it titled by
`op.name`, which is what it becomes.

| contract | real | oracle_engine | oracle_name | oracle_subject | oracle_inflect | null | dup5 | shot7 | shot45 | shot90 | wrongkind | dupmove | guards |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| shipped | 97 | 167 | **158** ✗ | 95 | 142 | 64 | 64 | 64 | 64 | 64 | 65 | **148** ✓ | fails `oracle_name` |
| shipped + name-titled oracle | 97 | **158** ✗ | 158 | 95 | 142 | 64 | 64 | 64 | 64 | 64 | 65 | 148 | *why oracle.py must move with the grader* |
| iteration 2 (`grade_v3`) | 114 | **166** ✗ | 167 | 138 | 159 | 64 | 64 | 64 | 64 | 64 | 69 | **167** ✗ | fails 2 |
| **iteration 3 (`grade_v4`)** | **114** | **167** | **167** | 137 | 159 | **64** | **64** | **64** | **64** | **64** | 69 | **152** | **PASS** |
| iter. 3, cancel any-word | 112 | 164 | 164 | 134 | 157 | 64 | 61 | 61 | 61 | 61 | 66 | 149 | fails 2 |
| iter. 3, cancel ≥ 0.5 | 112 | 166 | 166 | 135 | 158 | 64 | 63 | 63 | 63 | 63 | 68 | 151 | fails 2 |
| iter. 3, no stale rule | 114 | 167 | 167 | 138 | 159 | 64 | 64 | 64 | 64 | 64 | 69 | **167** ✗ | fails `dupmove` |
| iter. 3, kind-filtered pool | 114 | 167 | 167 | 137 | 159 | 64 | 64 | 64 | 64 | 64 | 69 | 152 | PASS — identical: kind-first is score-neutral |
| iter. 3, overlap before kind | 114 | 167 | 167 | 137 | 159 | 64 | 64 | 64 | 64 | 64 | **64** | 152 | PASS — rejected, see rule 2 |
| iter. 3, title-only brake | 114 | 167 | 167 | 137 | 159 | 64 | 64 | 64 | 64 | 64 | 69 | 152 | PASS — identical |

`dupmove` = 152 = 167 − 15 exactly: every double-booked move fails and nothing else does.

**Op-level failure profile on the real capture** (190 ops):

| bucket | shipped | iter. 2 | iter. 3 |
|---|---|---|---|
| ok | 56 | 76 | 76 |
| not found | 50 | 26 | **18** |
| wrong kind | — | — | **8** (was inside "not found") |
| wrong day | 14 | 30 | 29 |
| count: too many | 11 | 0 | 0 |
| stale copy on move | — | — | 1 |
| cancel residue | 3 | 2 | 2 |
| over-acted | 2 | 2 | 2 |
| **failing ops** | **80** | **60** | **60** |

Iteration 3 changes **no email verdict** on the real capture against iteration 2 (0 flips).
Its work is on the guards and the diagnosis.

**Read the `wrong day` rise honestly.** Roughly half the recovered work turns out to be on the
wrong date. That is not the contract failing — it is date error that was always there and
invisible, because the grader never got past finding the object. The +17 is net of it.

### Audits on the real capture

**Both remaining cancel failures are genuine, and one was a false pass under the shipped grader.**
- `World_Cup_Cleat_Launch.manufacturing-kickoff-3` — *"please take my stage slot off the run of
  show"* (Design Lead). The model kept `"Design walk-through at WC reveal"`, description
  `"Design Lead speaking slot on stage at reveal"`. That is the slot. Iteration 1's verifier read
  this as a false negative from the object's title alone; the description settles it. The shipped
  grader passed it because no object contained the literal phrase.
- `Partnership-with-deeptech-companies.boston-dynamics-cancel` — *"Boston Dynamics had to cancel
  our visit."* The model kept `"Boston trip - Boston Dynamics visit"`. In this capture the trip is
  split into separate events, so this is not the G-7 granularity case from the historical logs; it
  is a missed deletion.

**The one stale flag is a mis-diagnosis inside an email that fails anyway.**
`World_Cup_Cleat_Launch.reveal-event-date-and-venue-3`, `move 'Reveal Event'` — `event` is a stop
word, so the keyword set is `{reveal}` alone and every same-kind object with "reveal" in it
scores 1.0, including the design walk-through. The email's other op (`Reveal Rehearsal`) is on
the wrong day, so the verdict is unchanged; the reason string is wrong. This is the
single-content-word weakness (24 of 134 ops; VERIFY-phase2 §4.1) reaching the stale rule. It is a
corpus-naming problem — see open issue 1.

**The 8 wrong-kind claims.** Six are genuine disagreements between the model and the key
(`Team pizza party` keyed `todo`, the model made an event — K-7 says the key is wrong; `event day!`
keyed `todo`, the model made `"Event Pop Up"`; four `event` keys the model filed as to-dos:
*ask about patent overlap*, *retention conversation*, *approve trophy correction*, *podcast
taping*). Two are weak cross-kind claims of a sibling's object at low overlap (`Manufacturing
Kickoff`, `Pitch Breifing`) — correctly failing, loosely described.

**The `wrongkind` guard's +5 over the floor** (69 vs 64) is five emails where a *right-kind*
object of a sibling obligation shares a word with the obligation and sits on the same day. Under
kind-first ordering it is preferred to the wrong-kind object and passes the date check. The
shipped grader shows the same mechanism at +1. Overlap-first ordering takes this to exactly 64 and
was rejected for the reason in rule 2. All five involve a single-word keyword set.

**Stop-word sensitivity** (open issue 6 of iteration 2, now measured). Leave-one-out over all 97
words against `real`, `oracle_name`, `oracle_subject`, `oracle_inflect`: 12 words move some world,
never by more than ±2, and **none moves `oracle_name`**. 66 of the 97 never occur in any op name.
The list is insensitive; it is kept verbatim for the implementation.

---

## Closing the four blockers of iteration 2

| blocker | resolution |
|---|---|
| kind filter unaddressed | rule 2: cross-kind pool, kind-first tie-break, `wrong kind` reason. Score-neutral; closes the kind half of G-9 |
| `cancel` bypasses the assignment | rule 5: claims from what create/move leave unclaimed, at full overlap, either kind |
| `sb/oracle.py:52` titles by `match` | becomes `op.name.replace('_', ' ')`, in the **same commit** as the grader. Measured through the engine: 167. Left alone (titled by `match`), the gate reads 163 under this contract and 155 under iteration 2; changed alone, it reads 158 under the shipped grader |
| `sb/tests/test_e2e.py:58` would flip | it **does not flip**: the double-booked reschedule still fails, now under rule 4 with reason `moved, but 1 stale copy left behind`. The test will pin that reason so the mechanism is asserted, not just the verdict |

And two the blockers exposed:

- `oracle_engine` 166 in iteration 2 was **not** a corpus ambiguity after all. It was the dead
  `email_id` tie-break on the engine path (rule 2). With turn membership as the tie-break the
  `Company-Retreat` name nesting resolves itself: the fresh object is claimed, the inherited one
  is the sibling's. The G-5 nesting is still real and still wants a lint (open issue 1); it just
  does not fail the oracle any more.
- The synthetic agent was a double-booker (see the guard set). Corrected numbers are above.

---

## Open issues

1. **Single-content-word obligations** (24 of 134, e.g. `Reveal Event` → `{reveal}`, `board`,
   `budget`, `pitch`) are the soft spot of every rule here: the identity rule's only defence on
   them is the date predicate, the stale rule can mis-describe a sibling as a copy, and the
   `wrongkind` guard sits 5 above the floor because of them. The G-5 **name-aware lint** — flag a
   same-kind obligation whose keyword set is a subset of a sibling's, and any obligation whose
   keyword set is one word — belongs to phase 5 alongside the corpus edits it will demand.
   `sb/schema.py` lint #5 still checks the old `match`-collision rule; harmless but dead once
   `match` is unread. Replacing it with the name-aware rule fails the current corpus on 10 names,
   so it cannot ride in the grader commit.
2. **`event day!`** (`Sponsoring-Marathon.approval-of-event`) is the only op whose name is
   entirely stop-words; the fallback makes `{event, day}` its keywords, words every other op
   stops. In this capture it claimed a wrong-kind `"Event Pop Up"`. A live gaming surface if the
   corpus ever grows another such name; rename in phase 5.
3. **Descriptions are in the identity haystack** (G-3). Iteration 1's verifier confirmed the
   real-capture flips that depend on it are legitimate, and the cancel audit above shows a case
   where the description is what makes the grader right. Kept, deliberately. The model is still
   not told; that is phase 3's prompt work, not the grader's.
4. **K-7's four mis-keyed kinds** now surface as `wrong kind` instead of `not found`. Phase 5.
5. **One capture, one model, one seed, one lever set.** The guards are model-independent; the +17
   is not. It needs a second captured run before anyone quotes it.
6. **Everything here is grader-versus-grader.** No measurement in this document says the score
   moved toward *truth*. Only the phase 1d hand-grade can say that, and it is deferred behind
   C-10. Per the register's status legend this can reach `fix proposed` — never `verified`.

---

## Implementation notes

- `grade_email(answer, ctx, state, turn)` keeps its signature. No email id is needed.
- `EmailResult`/`details` keep their shape (`passed, label, expected, actual, reason`);
  `sb/live/runner.py:_print_email`, the capture's `verdicts`, and `sb/regrade.py` read them
  unchanged. `sb/analyze.py` parses only `PASS`/`FAIL` lines.
- New reason strings: `wrong kind: created a <kind>, expected a <kind>`; `over-created: N
  equally-matching <kind>s for one obligation`; `moved, but N stale copy left behind
  (double-booked)`; `should be cancelled, but still on the calendar`.
- `sb/oracle.py` titles by `op.name.replace('_', ' ')`. `Store.find_in_node` substring-matches
  the title, which is exact for the oracle's own objects.
- `sb/tests/test_e2e.py:58` keeps `assert not passed` and gains an assertion on the reason.
- The guard harness becomes `sb/tests/test_grader_guards.py` so no future change can quietly
  inflate the score. It runs in seconds.
- `webapp/api/_lib/sb/` is regenerated at build by `webapp/scripts/vendor_sb.py`; nothing to do.
