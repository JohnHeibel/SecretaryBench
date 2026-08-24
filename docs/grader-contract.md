# The grader's identity contract — phase 2

**Status: `fix proposed`.** Designed and measured, **not implemented**. `sb/grader.py` is
unchanged. One earlier design was rejected outright by an adversarial pass; this is the
second iteration and it has not yet been re-verified.

Register: **G-1, G-2, G-5, G-7, G-8**, and **C-10**. Evidence: `captures/baseline-sonnet-4-5`.

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
by the obligation's own name scores **148/167** under today's grader. The benchmark cannot
award full marks to a flawless assistant.

---

## The proposed contract

Three changes. They must land together — measured separately, each looks like it does nothing,
because fixing one relocates failures into another.

### 1. Identity from the obligation's name, not a generated keyword

Keywords are the **stemmed content words of `op.name`** (stop-words dropped; falls back to all
words when a name is entirely stop-words, e.g. `event day!`). Matching is word-level against
`title + description`.

Consequences: `match` is no longer consulted, so **`fix_match.py` becomes unnecessary and the
corpus needs no mutation** — which resolves the largest source of the C-10 fork. Stemming is
required, not cosmetic: without it the rule is *less* inflection-robust than the substring rule
it replaces (inflected-title oracle 68 vs 127).

### 2. Exclusive assignment replaces the exactly-one rule

Every (obligation, object) pair is scored by the fraction of the obligation's keywords present.
Pairs are sorted and greedily assigned so **each object serves at most one obligation and each
obligation claims at most one object**. Tie-breaks, in order: overlap score, kind match, then
attribution to the email being graded.

This eliminates the fake-duplicate category outright (11 → 0 on the real capture) without
letting a sibling's object satisfy the wrong obligation.

### 3. A volume brake, scoped to the turn

After assignment, an obligation **fails** if another object created **in the same turn** matches
it at least as well and was claimed by nobody. That is a real duplicate. Objects inherited from
earlier emails are untouched — they are their own obligation's answer, not over-creation.

Scoping to the turn delta rather than to `email_id` matters twice: it works on every code path
(the engine path has no email id), and it **cannot be defeated by a mis-stamped `email_id`**
(register A-5).

Note for implementation: `_turn_delta` constructs fresh `Obj` instances, so turn membership
must be tested **by value** — `(kind, title, when, email_id)` — never by `id()`.

### Explicitly not included

**No `dateok` term in the tie-break.** An earlier design preferred objects whose date already
matched, which turned grading into an active search for something that fits. It is free to
remove on every measured world and it removed 63 points of gaming headroom.

---

## The guard set

The first iteration passed three guards and was still broken, because the guards could not vary
the one thing the contract stopped constraining: **how many objects the model creates**. These
adversaries are built from the ones that broke it.

| guard | what it is | requirement |
|---|---|---|
| `oracle_engine` | `sb/oracle.py` through `sb.engine.run` | **167** — `sb/scale.py:127` gates on it |
| `oracle_name` | perfect agent, titles = obligation name | **167** |
| `oracle_subject` | perfect agent, titles = email subject | high — realistic naming |
| `oracle_inflect` | perfect agent, pluralised titles | high — inflection robustness |
| `null` | creates nothing at all | **64** exactly (register V-3) |
| `dup5` | right answer, then 5 copies of every object | **must not exceed null** |
| `shot7/45/90` | never reads a date; one object per day over N days | **must not exceed null** |
| `wrongkind` | right title and date, wrong kind | informational |

`null` and a uniform `+3 day` shift are **structurally incapable** of separating two contracts on
this corpus (56 no-action + 8 all-cancel emails; 112 `eq` / 12 `by` / `exact_day` on 134 of 134
ops). They are retained as regression tripwires, not as evidence.

---

## Measured

`sb/grader.py` unchanged; contract applied via a harness over the certified capture.

| world | shipped | proposed | note |
|---|---|---|---|
| **real** (`claude-sonnet-4-5`, certified) | 97 | **114** | +17 |
| `oracle_engine` | 167 | **166** | one corpus ambiguity, below |
| `oracle_name` | **148** ✗ | **167** ✓ | shipped cannot award full marks |
| `oracle_subject` | 92 | **139** | realistic naming, +47 |
| `oracle_inflect` | 132 | **159** | +27 |
| `null` | 64 | **64** ✓ | |
| `dup5` | 64 | **64** ✓ | was **167** before the brake |
| `shot7` | 64 | **64** ✓ | was 93 |
| `shot45` | 64 | **64** ✓ | was 128 |
| `shot90` | 64 | **64** ✓ | was 148 |
| `wrongkind` | 65 | 69 | +4, see open issues |

Failure profile on the real capture: `not found` 50→26, `wrong day` 14→30,
`count: too many` 11→**0**, cancel residue 3→4, over-acted 2→2.

**Read the `wrong day` rise honestly.** Roughly half the recovered work turns out to be on the
wrong date. That is not the contract failing — it is date error that was always there and
invisible, because the grader never got past finding the object. The +17 is net of it.

---

## Open issues, all real

1. **`oracle_engine` is 166, not 167**, so `sb.scale`'s mandatory gate would fail.
   The cause is a **corpus ambiguity, not a contract flaw**: in `Company-Retreat`, the
   obligations `'Retreat Company Meeting Call'` and `'Company Retreat'` both reduce to
   `{company, retreat}` — lexically indistinguishable once stop-words are removed. Register
   **G-5** predicts exactly this and its name-aware lint variant flags 10 such cases. Fix
   belongs to phase 5, or the gate needs a documented exception.
2. **`sb/oracle.py` must change in the same change-set.** It titles by `op.match`
   (`sb/oracle.py:51`); the contract keys on `op.name`. Left alone, the gate reads 155.
   Per `CLAUDE.md` this must not ride in the same commit as a corpus edit.
3. **`sb/tests/test_e2e.py:58` will flip.** It asserts a double-booked reschedule fails. The
   volume brake should still catch that case, but the test must be re-baselined deliberately
   and in writing, not silently.
4. **The kind filter is unaddressed.** Kind mismatch rises to a larger share of the residual,
   and `wrongkind` gains 4 over the null floor. The register's own entry condition — identity,
   count, cancel **and kind** as one contract — is therefore not yet met.
5. **`cancel` needs review.** It currently uses a separate full-overlap test rather than going
   through the assignment. G-7's option list recommends unifying them.
6. **The stop-word list is hand-written** and long. Its sensitivity has not been measured.

---

## What would make this trustworthy

- **One capture, one model, one seed, one lever set.** The guards are model-independent; the
  +17 is not. It needs a second captured run before anyone quotes it.
- **Everything here is grader-versus-grader.** No measurement in this document says the score
  moved toward *truth*. Only the phase 1d hand-grade can say that, and it is deferred behind
  C-10. Per the register's status legend this can reach `fix proposed` — never `verified`.
- The guard harness is a scratchpad prototype. It must become a real test suite in
  `sb/tests/` so no future change can quietly inflate the benchmark.
