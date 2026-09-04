# Phase 1d hand-grade — running findings

Cross-cutting observations from the hand-grade of `captures/baseline-sonnet-4-5`, kept as
the pass proceeds. Per-item verdicts live in `audit-baseline-sonnet-4-5.md`; this file is
for what the items have in common.

Every number here is reproduced by:

    .venv/bin/python docs/_repair/handgrade_probes.py

**Status:** in progress. 11 of 60 Part A items verdicted, 1 left open (item 6), Part B not
started. Nothing here is a register finding yet; candidates are marked and the counts that
would promote them are given. No corpus or grader change: corpus sha `03e0d963b9866d8f`.

---

## Finding 1 — a key's identity words are often not in the email *(candidate: new register ID)*

**This is the title analogue of K-6** (`answers the prose does not pin`, which covers dates
only). No existing finding covers it: G-1 and G-8 were about the *matching rule* and both
dissolved in phase 2; G-5 compares obligations to their *siblings*, not to the email.

Measured over all 125 `create`/`move` ops (probe 3):

| | count | share |
|---|---|---|
| at least one identity word absent from body + subject | **30** | 24% |
| every identity word absent | **2** | 1.6% |

Both counts are a **floor**. The probe reads the *raw* body, which still carries anchor token
names (`{!launchmeeting = this:THU}`) that the model never sees, so it credits the model with
words a rendered email does not contain.

The two categorically unanswerable ones:

- `AI_Final_Meeting_Review` → `{ai}` (worksheet item 3, also **PLAN**-failed on the date)
- `board_signoff` → `{board, signoff}`, `Planning.can-t-do-thursday`

The clearest illustration is **worksheet item 9**, `Friday_AI_Review` → `{ai, friday}`. The
entire email is *"Let's meet Friday, June 19th, 2026. Im packed the rest of the day."* The
token `ai` is nowhere in it; it comes from the node name `Enterprise_Ai_Selection`, which the
model never sees as a title cue. The model produced `"Meeting with CTO"` on the right day, of
the right kind, correctly stamped, and scored zero identity overlap.

**Why this is a grader finding and not only a corpus one.** `docs/grader-contract.md:174`
already pins `oracle_subject` at **137/167**: a *flawless* agent that titles objects with the
email's subject line loses 30 emails to titling alone. So the cost of the title channel to a
correct agent is measured and large, independent of this pass.

**What it does not yet say.** A missing word is not automatically a failure — whether the op
can still pass depends on the contract's overlap threshold and the remaining words. Turning
"30 keys demand an absent word" into "N keys are unpassable" needs the threshold applied per
op, which has not been done.

**Promote when:** the pass is complete and the `GRADER` tick count is known. That count is
the false-negative rate, which is the number phase 2 exists to improve and nobody has.

Related, same probe: `Pitch Breifing` → `{breif, pitch}` demands a **misspelled** token
(`Sponsoring-Marathon.pitch-deck-2`). An authoring typo that no correct title can satisfy.
Belongs with K-8's name hygiene.

---

## Finding 2 — the kind convention, and six mis-keys where K-7 named four

The corpus has a consistent latent rule for `event` vs `todo`, and it matches the only
documented one (`sb/live/runner.py:89`: to-dos are "tasks that have a deadline"):

| probe | keyed as expected | outliers |
|---|---|---|
| verb-named ops (actions) | 25 `todo` | **4** `event` |
| gathering nouns | 34 `event` | **3** `todo` |

Discounting `thank_you_lunch` (a genuine gathering the verb probe false-positives on) and the
three gathering-nouns keyed `todo` that are really verb-named actions about a meeting, six ops
are mis-keyed against the corpus's own convention. **K-7 named four** and said its conservative
detector would miss ambiguous cases; it did.

| op | keyed | should be | K-7? | item |
|---|---|---|---|---|
| `Team_pizza_party` | todo | event | yes | 2 |
| `AI_Final_Meeting_Review` | todo | event | yes | 3 |
| `approve_trophy_correction` | event | todo | yes | 7 |
| `Contact People Added To List` | event | todo | yes | — |
| `event day!` | todo | event | **no** | 1 |
| `ask_about_patent_overlap` | event | todo | **no** | 4 |

**A hypothesis was tested and falsified.** "Gatherings with a proposed but unconfirmed time
become to-dos" would have made the pizza key internally consistent. Probe 2 finds **10**
event-keyed creates sitting on emails that hedge as hard or harder (`Proposed` ×2,
`tentativ` ×2, `Pencil`, `I think`), `project_atlas.launch-dinner` ("Penciling it in") among
them. No rule separates them because there is no rule: kind selection is documented in exactly
one line, and nothing in `docs/`, `corpus/`, the ADR or either design doc adds to it.

**Cascade.** `_wire_obligations` (`sb/schema.py:393`) makes a sibling `move` inherit the
create's kind, so the pizza mis-key fails **two** emails for one authoring error.

---

## Method note — `GRADER` vs `KEY`, and why the split matters

The two codes route to different phases, so they must not be merged:

- **`KEY`** — the obligation itself is unsupported (wrong kind, wrong date, an obligation
  nobody would infer). Fix is a corpus edit. **Phase 5.**
- **`GRADER`** — the obligation is right and the model satisfied it, but the grader could not
  see that. Fix is the identity contract. **Phase 2/3.**

Item 9 is `GRADER`: there *is* a meeting on Jun 19 and the model booked it. Item 2 is `KEY`:
the obligation says `todo` and it should say `event`. Coding item 9 as `KEY` would send an
identity-contract defect to the rename pile and undercount the false-negative rate.

`[close]` is used in notes, not as a code, to mark an op where the model's alternative was
defensible even though the tick is `MODEL` (item 5). Phase 5 wants that list: those are the
obligations whose names or prose need sharpening, and it is a different list from K-7's.

---

## Which existing register findings are actually biting

Instances hand-confirmed so far, as a check on whether phase A's estimates hold up:

| finding | what it predicts | confirmed here |
|---|---|---|
| **K-1** emails served after a date their own body states (18) | no answer can be right | items **3**, **10** |
| **K-6** exact-day keys with no cue in the prose (5) | model penalised for a date nothing pins | items **7**, **12** — both `Innovation-comp`, 2 of the 3 K-6 names by construction |
| **K-7** kind contradicts the obligation (4 named) | right object, wrong column | items **1**, **2**, **4**, **7** — two of them *not* among K-7's four (finding 2) |
| **K-7** weekend `eq` dates (14 of 112) | implausible business dates | item **12** (Sun Jun 21 to review a render) |

Item 12 carries **two** defects at once, K-6 and the weekend half of K-7, which is worth
remembering when phase 5 counts ops to edit: the defect count and the op count differ.

---

## Running tally (Part A, 11 of 60)

| code | items | n |
|---|---|---|
| `KEY` | 1, 2, 4, 7, 12 | 5 |
| `PLAN` | 3, 10 | 2 |
| `MODEL` | 5, 11 | 2 |
| `GRADER` | 9 | 1 |
| open | 6 | 1 |

**Do not read a ratio off this yet.** Eleven items ordered by serve date is not a sample, and
`Innovation-comp` — which K already flagged for dominance — supplies five of them. The number
that matters is the `GRADER` count at the end of the pass.

Two clean `MODEL` items so far and both are unambiguous: item 11 created nothing at all, item 5
filed a gathering as a weekend to-do. Neither is a near miss, which matters — it means the
`MODEL` bucket is so far measuring real failures rather than grader strictness.
