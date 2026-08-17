# VERIFY — red-team of the phase-A merge

Adversarial pass over `docs/benchmark-repair.md` against the code, the corpus and the
four recorded logs. **This file reports; it changes nothing.** No register edit, no code,
no corpus, no live run. `scripts/recover_corpus.py` and `scripts/fix_match.py` were read,
never executed.

Every objection below carries a command, a `file:line`, or a quoted line. Scratch scripts
live outside the repo (`/private/tmp/.../scratchpad/v_*.py`); every number here was
produced by a script written for this pass, not by re-running a category agent's method.
`PYTHONPATH=/Users/jamesoc/dev/SecretaryBench` prefixes each command.

**Bottom line.** ~200 discrete claims checked: 6 headline numbers re-derived from scratch,
~50 further numeric claims re-derived, ~145 `file:line` anchors opened, 8 grep assertions
re-run. **The register's arithmetic is unusually sound** — the headline numbers reproduce to
the digit, no anchor is fabricated, no severity drifted, all 50 IDs survive. The failures are
almost entirely in *interpretation*: sentences that assert a cause, a proof, or a permanence
the measurement underneath does not carry. Two are substantively wrong, eight are overstated,
and **one of the two wrong ones is the sole justification for making phase 1 "time-critical"** —
the change the register calls "the most consequential thing in this document".

| verdict | count | the ones that matter |
|---|---|---|
| CONFIRMED | ~180 | all six headline numbers; G-4 and G-1 exact; 2 of the 3 judgment calls |
| WRONG | 2 substantive + 9 minor | hazard box "permanently destroys"; C-1's haiku levers; 6 off-by-N anchors |
| OVERSTATED | 8 | O-1 "can ever be re-graded"; resolution 2's "proves"; K-2's title; G-1's prompt claim |
| UNFALSIFIABLE | 3 | G-3's writing-style clause; "785 feasible" sans grid; G-6's "pool depth" |

---

## 1. The six headline numbers

### G-4 — the oracle titling by email subject scores 92/167 · **CONFIRMED, exactly**

Written from scratch: a scheduling-perfect agent that reuses only the *date* arithmetic of
`sb/oracle.py` (`_target`, `_as_dt` — the half stipulated perfect) and chooses titles by a
policy under test, driven through the shipped `sb.engine.run` + `sb.grader`.

```
.venv/bin/python scratchpad/v_g4.py     # corpus/, seed 42, start 2026-06-01, Levers(1,5,7)
```

| policy | my score | register | action-only (n=111) |
|---|---|---|---|
| P0 `" ".join(op.match)` | 167/167 | 167/167 | 111/111 |
| P1 `op.name` verbatim | 160/167 | 160/167 | 104/111 |
| P2 `op.name` humanized | 157/167 | 157/167 | 101/111 |
| **P3 email subject** | **92/167 (55%)** | **92/167 (55%)** | **36/111** |
| P4 subject + humanized name | 140/167 | 140/167 | 84/111 |
| P0 + realistic description | 141/167 | 141/167 | 85/111 |

Every cell matches. The claim is also **robust to the one ambiguity in "scheduling-perfect"**:
"scheduling-perfect" does not say how the agent finds the object it created earlier. My
`LOOKUP` variant resolves prior objects by title substring exactly as `sb/oracle.py:55,60`
does, and gives 92/167. A stricter `IDENT` variant, where the agent has perfect internal
memory of which store object serves which obligation so that *only the title string* varies,
gives **95/167 (57%)** — still inside the 54–59% band. The load-bearing claim survives both
readings.

**One nuance the register's prose does not draw out, though its own table contains it.**
P3 scores **36/111 on the acting half, below all three real models** (opus 39, sonnet 40,
haiku 42 of 111 — reproduced below). The subject-titling oracle lands "inside the band" only
because it collects the same 56 free no-action points every model collects. The defensible
statement is therefore stronger than "the band is reachable with zero scheduling errors": on
the emails that actually require an action, a *scheduling-perfect* agent under a plausible
title policy is beaten by every real model. That sharpens G-4; it does not weaken it.

### G-1 — 32/125 ops demand an undiscoverable keyword; 8% vs 48% · **CONFIRMED, exactly**

```
.venv/bin/python scratchpad/v_g1.py
```

- create/move ops = **125** (108 create + 17 move; 9 cancel; 134 total).
- Keyword absent from the node's rendered mail (haystack = rendered bodies + subjects):
  **32/125**. Body-only haystack gives 37; adding sender/recipients changes nothing.
- Pooled over the three 167-corpus runs at op level: present **146/303 = 48%**,
  absent **8/99 = 8%**. Per model, present: opus 43%, sonnet 50%, haiku 51% — the
  register's `(43% / 50% / 51%)` to the digit.

**I ran a control the register did not, and it survives.** The obvious confound is that the
absent-keyword ops cluster in the two nodes no model can do (K-4). They partly do —
`Sponsoring-Marathon` 4/7, `World_Cup_Cleat_Launch` 7/16, `pizza-party` 4/4. Restricting to
the nine nodes that contain *both* classes and pooling within them:

```
present 74/201 = 37%   absent 5/84 = 6%
```

The 40-point gap is not a node effect. G-1 is the best-controlled finding in the fan-out
after G-4.

### A-1 — day-scoped re-grade gives opus 55, sonnet 57, haiku 43 · **CONFIRMED, exactly**

```
.venv/bin/python scratchpad/v_a1.py
```

| run | raw | no-action passes | acting passes | nulled | day-scoped | emails/day |
|---|---|---|---|---|---|---|
| opus | 90 | 51/56 | 39/111 | 35 | **55** | 2.9 |
| sonnet | 91 | 51/56 | 40/111 | 34 | **57** | 2.9 |
| haiku | 98 | 56/56 | 42/111 | 55 | **43** | 10.4 |

56–57% of every headline score is the abstain check (51/90, 51/91, 56/98) — confirmed.
`35 of 51 / 34 of 51 / 55 of 56` reproduces the figure the register uses in Contradiction 1.

### V-1 — 43,474 chars, 18,389-char gap; V-3 — 64/167 floor · **CONFIRMED, exactly**

```
.venv/bin/python scratchpad/v_v1.py ; scratchpad/v_v1b.py ; scratchpad/v_v1c.py
```

- **43,474** = raw authored body + subject + sender, summed over 167 emails. Exact, and
  lever-independent. (Rendered bodies alone are 38,472; +subject 42,949.)
- Mean authored body **227.5** chars — the register's 227.
- **18,389** = raw body chars over the largest needle window
  (`Day-of-execution_and_Aftermath.launch-livestream` → `.thank-the-team`, email-span 83),
  excluding the setup email, including the payoff. Exact under that convention; the four
  plausible boundary conventions give 18,193 / 18,382 / **18,389** / 18,578.
- Null model (never calls a tool): **64/167 = 38.3%**, and the passing set is exactly
  {56 no-action} ∪ {8 cancel-only} — verified by set equality, not by assertion.
- V-3's table reproduces: on-floor 56/64, 55/64, 60/64; actionable 34/103 = 33.0%,
  36/103 = 35.0%, 38/103 = 36.9%.

### K-2 — 87 of 127 resolved answer dates move · **CONFIRMED, exactly**

```
.venv/bin/python scratchpad/v_k2.py ; scratchpad/v_k1.py
```

- Resolved answer dates n = **127**, moved between `daily_max=5` and `21`: **87**.
- Serve dates moved: **160/167**. Days 57 → 16.
- I independently reproduced the whole **past-dated-ops column** of K-2's table:
  `daily_max` 5 → **11**, 8 → **16**, 13 → **11**, 21 → **4**, 30 → **2**. Exact, all five.
- 19 of 100 seeds raise `InfeasibleSchedule` — confirmed at `n_days` 100 and 400.
  *Sharpening:* at the runner's shipped default `--days 60` it is **31 of 100**.

### C-1 — levers recovered uniquely out of 785 · **CONFIRMED for opus/sonnet, WRONG for haiku**

```
.venv/bin/python scratchpad/v_c1.py    # my grid:  dmin 1-5 x dmax 1-30 x uh 1-30
.venv/bin/python scratchpad/v_c1b.py   # C.md's grid: dmin 1-5 x dmax 1-29 x uh in {3,5,7,10,14,21}
```

- **opus and sonnet: unique at `(1, 5, 7)`.** This holds in C.md's 785-combination grid *and*
  in my much wider 4,082-combination grid. The identification is stronger than claimed.
- **haiku: not three combinations — seven.** `urgency_horizon ∈ {1,2,3,4,5,6,7}` all
  reproduce haiku's 16-day plan exactly, with `daily_min=1, daily_max=21`. C.md's "three"
  is an artifact of a grid that samples only `{3,5,7,10,14,21}`; inside that grid it is
  correct, and my run of that exact grid returns `{3,5,7}`. The honest statement is
  **"`urgency_horizon ≤ 7` is unidentifiable for haiku"**.
- **The grid was dropped in the merge, and it is the only thing that makes "785" checkable.**
  `C.md:56` states it: "`daily_min` 1-5 × `daily_max` 1-29 × `urgency_horizon ∈ {3,5,7,10,14,21}`".
  The register (`benchmark-repair.md:765`, `:343`, correction 15 at `:1457-1460`) carries the
  number and not the grid. As written, "one of 785 feasible combinations" is unfalsifiable and
  "three values (3, 5, 7) reproduce it exactly" reads as a claim about identifiability that is
  false. **Fix: restore the grid to the register, and restate haiku as `uh ≤ 7`.**

---

## 2. The three synthesizer judgment calls

### 2a. Withdrawal of O-2's opus trace-loss claim · **the withdrawal is CORRECT; the replacement claim is OVERSTATED**

**The withdrawal itself holds.** `sb/grader.py:151-152` builds `title_set` from the node's
cumulative pool, and `:168` renders only that set, so a title's first appearance in an
`actual` field bounds its creation day from *above*, not equals it. O's per-day deficit
measure is therefore invalid, exactly as the register argues. O's aggregate numbers
reproduce (`grep -o 'create_event\|create_todo' | wc -l`): **89 / 41 / 18 / 20** create_*
calls — matching `O.md:131-134` — against distinct `actual` titles of **69 / 70 / 67 / 65**
by my regex (O counted 71/70/68/65; a 2-title parse difference in two logs). Every
comparison keeps its direction: opus 89 ≥ 69 proves nothing either way; sonnet 41 < 70 and
haiku 18 < 67 prove loss. The withdrawal is right, and V-2's opus conclusion does not fall.

**But the sentence that replaces it does not follow.** `benchmark-repair.md:1235-1236`:

> "Retained can never *exceed* actual, so equality on 56 of 57 days proves no `get_email`
> was dropped on those days."

Retained = the known lower bound (one `get_email` per email) is consistent with *any*
amount of loss above that bound: actual could be 7 on a 5-email day with 2 lost. Equality
with a lower bound is not a proof of zero loss. And `:1237-1239`:

> "count equality demonstrates that opus **serialised** its tool calls (one content block
> per assistant message)"

is contradicted by O's own synthetic case 2 (`O.md:110`): five blocks in **one** event under
one id return all five. Count equality is equally consistent with lossless batching, so it
demonstrates nothing about serialisation — and serialisation is explicitly the step the
register says is load-bearing for excluding a hidden `search_inbox` (`:1240-1241`).

The register concedes the real dependency 30 lines later (`:1267-1268`: "whether the current
CLI emits one `assistant` event per content block … neither confirmed it"), which makes the
body text and the open-questions bullet disagree with each other.

**Defensible version:** *opus's retained trace meets the known lower bound on 56 of 57 days
and tracks batch size across all five batch sizes, which is what a lossless trace looks like
and is inconsistent with sonnet's `{1: 57}` and haiku's 40-for-167. Whether opus lost a
`search_inbox` is not settleable from the log.* Consequently **correction #8 (`:1437`) should
read "a tight lower bound for opus", not "a measurement for opus"**, and V-2's title
("real for opus and unmeasurable for sonnet and haiku") should say "tightly bounded for opus".

Verified inputs, all reproduced: opus `get_email` per-day histogram `{1:11, 2:13, 3:12, 4:12,
5:9}` = 166 total; sonnet `{1: 57}` = 57; haiku 40. `search_inbox` on **1/57, 1/57, 0/16,
1/19** days, and for both opus and sonnet the one day is day 1.

### 2b. A-1's rank inversion ruled a lever artifact · **CONFIRMED — and I can make it causal, not correlational**

The register's argument is correlational as stated ("55 of haiku's 56 … sit on a day with a
visible `create_*`"): it uses the model's own observed behaviour, so it cannot distinguish
"the lever forced it" from "haiku happened to act on those days".

```
.venv/bin/python scratchpad/v_a1b.py
```

The model-independent version settles it. Under a day-scoped rule, a no-action email can
only survive if *nothing* was created that day; so the ceiling is fixed by the serve plan
alone — count the no-action emails that share their day with **no ops-carrying email**:

| `daily_max` | days | no-action emails alone on their day | ceiling on day-scoped no-action passes |
|---|---|---|---|
| **5** (opus, sonnet) | 57 | **10 / 56** | 10 |
| 8 | 37 | 3 / 56 | 3 |
| 13 | 21 | 0 / 56 | 0 |
| **21** (haiku) | 16 | **1 / 56** | 1 |
| 30 | 11 | 0 / 56 | 0 |

Haiku's day-scoped score is capped **nine points below** opus's and sonnet's before any model
behaviour is considered. The observed spread in surviving no-action passes (opus 16, sonnet
17, haiku 1) is mostly that ceiling. The confound is structural, not incidental, and the
register's call — *A-1's contribution is the weight of the abstain check, not the inversion* —
is right. It is also right in the direction of the trace-loss caveat: haiku has the lossiest
trace (18 retained creates) and therefore the most opportunity for a *spurious* surviving
pass, yet has the fewest, because the ceiling binds.

**Recommendation:** replace the register's sentence at `:1189-1192` with the ceiling table.
It is the same conclusion with a causal argument instead of a correlational one.

### 2c. The phase re-ordering · **the argument does NOT hold as written — both of its load-bearing premises are overstated**

The re-ordering rests on two claims. Both are individually defensible in spirit and
overstated in text, and one is simply wrong.

**(i) O-1: "No run already paid for can ever be re-graded, at any price short of running it
again."** (`:82-84`, restated at `:632`.) **OVERSTATED, and contradicted three times inside
the register itself.**

- `:382-383` (G-2): "Under a rule of 'at least one matching object on an expected day',
  opus 90→97, sonnet 91→98, haiku 98→109, sonnet176 102→111."
- `:526` (A-1): "Re-grading the same logs day-scoped … gives opus 90→55".
- `:614-615` (A-6): "Counterfactuals: object never existed → 90→91; no-action re-evaluated at
  end of run → 90→89."

Each of those *is* a re-grade of a paid run. I reproduced two of them from the logs alone
(`scratchpad/v_g2.py`: opus 90 → **97**, matching exactly; `scratchpad/v_a1.py`:
55 / 57 / 43, matching exactly). The `actual` field carries every matched object's title,
date, time and duration, which is enough to re-score any rule that *relaxes or re-partitions*
the already-rendered `title_set`.

**Defensible version:** *no run can be re-graded under a rule that would newly admit objects
the log never rendered (anything that widens `_title_hit`, changes `op.kind`, or re-attributes
by `email_id`), because `grader.py:151-152` filtered those out before printing. Rules that
relax the count check, re-scope the no-action delta, or drop an object can be evaluated
offline today, and the register does so three times.* That is a real and material limit — it
still blocks G-1, G-3 and G-8 from offline re-scoring — but it is not "nothing can ever be
re-graded", and the weaker version does not by itself move O from phase 1 to phase 6.

**(ii) C-1's deadline: "Either one run today permanently destroys the only provenance the
three surviving runs have"** (`:23-25`, hazard box) and "the recovery works only while the
corpus is frozen" (`:768`). **WRONG.**

```
git ls-files corpus/            -> 16 files, every node json tracked
git status --porcelain corpus/  -> empty (clean)
git diff 24331fb HEAD --stat -- corpus/  -> empty (unchanged since 24331fb)
git check-ignore -v corpus/nodes/pizza-party.json -> exit 1 (not ignored)
```

Both scripts do exactly what the register says — `scripts/recover_corpus.py:131,133` unlinks
and rewrites `corpus/nodes/*.json` and `:25` targets
`https://secretarybench.vercel.app/api/nodes`; `scripts/fix_match.py:86,88` unlinks and
rewrites in place. But every file they touch is **git-tracked, committed and clean**, so
`git checkout -- corpus/nodes` (or `git show 24331fb:corpus/nodes/<f>.json`) restores them
byte-for-byte. The lever recovery is reproducible from git at any future date. The register
knows the corpus is fully tracked — it relies on that fact in C-6 and in correction 19
(`git log --all -- corpus/`, `git grep` across every ref) — so this is an internal
inconsistency, not a missing fact.

**Consequence for the phase table.** Phase 1 is still *worth doing* (it is free, and a
committed sidecar beats a re-derivation), but it is **not time-critical**, and
"time-critical" is the word doing the work at `:110`. Neither premise survives at full
strength: (i) is half true, (ii) is false. The *conclusion* may still be right — G-4 alone
shows G iterates offline today without O, and that argument I confirmed independently — but
**the re-ordering should be re-argued on G-4 alone, and the "free, time-critical" label on
phase 1 should be downgraded to "free, do it early".** This is the single most important
error in the document, because `:47-49` tells the next session that this ordering "is the
most consequential thing in this document" and asks for sign-off on it.

---

## 3. Claims that assert more than the evidence carries

| # | where | verdict | the defensible version |
|---|---|---|---|
| 1 | `:914` K-2 title, "the corpus is **only** coherent at `daily_max=21`" | **OVERSTATED** — its own table shows `daily_max=30` strictly better (1/2/3 defects vs 1/4/6), and I reproduced the past-dated-ops column exactly (4 at 21, **2 at 30**) | "coherence improves monotonically above `daily_max=13`; 21 is the *documented* lever, not the uniquely coherent one" |
| 2 | `:99` and `:345`, "the largest needle gap ~7.8k tokens" | **OVERSTATED / unsourced** — the measured 18,389 body chars is 4.6k tokens at the 4.0 ch/token the register's other figures use; adding subject+sender over the same 83-email window gives ~5.4k. The step to 7.8k needs an unstated ~10k-char JSON envelope | "~5.4k tokens of mail, more with per-email JSON overhead" — the conclusion (two orders of magnitude under a 200k window) is untouched |
| 3 | `:397-398` G-3, "the spread tracks how verbose each model is, so the metric is partly measuring writing style" | **UNFALSIFIABLE from any artifact** — descriptions are never rendered (`grader.py:125-129`, and O-5 says so), so description verbosity cannot be measured from the logs. It also fails the one verbosity proxy that *does* exist: mean title length is opus 47 > haiku 40 > sonnet 36, while description-only matches are opus 23 > sonnet 14 > haiku 8 | drop the causal clause; keep "23/101, 14/97, 8/97, 2/91 objects matched through the description alone". What would settle it: O-3's state dump, which records descriptions |
| 4 | `:450-452` G-6, "0 prior … 42/81; 1–2 65/156; 3–5 33/111; 6+ 14/54" | **not reproducible from the register alone** — "pool depth" is undefined (prior *obligations* in the answer key, or prior *objects* the model made?). My prior-obligation definition gives the same 402 ops and the same 154 passes but buckets 42/78, 56/132, 38/117, 18/75 | the monotone decay (54→42→32→24 my way, 52→42→30→26 theirs) is robust; state the bucketing rule next to the numbers |
| 5 | `:382-383` G-2's counterfactual | **method-sensitive, method not carried** — `G.md:160-161` restricts to `eq`/`any_of` predicates "where the day is unambiguous". Without that restriction I get opus 97 (same), sonnet **100** (vs 98), haiku **110** (vs 109) | carry the restriction into the register; the number moves by 1–2 without it |
| 6 | `:1008` K-7, "fails **both** of that node's graded emails" | **OVERSTATED in the merge** — `pizza-party.json` has **4** graded emails (`end-of-year-pizza-party`, `pizza-place-selection`, `pizza-order-deadline`, `client-demo-conflict`); `K.md:621` correctly says "two of the node's four graded emails" | restore "two of the node's four" |
| 7 | `:367-368` G-1, "`SYSTEM_PROMPT` (`sb/live/runner.py:87-107`) says nothing about titles, descriptions, or **one-object-per-obligation**" | **OVERSTATED** — `runner.py:107` ends `"RESCHEDULING: to move an event, update_event it (delete-and-recreate is also fine). To cancel, delete it. **Never leave duplicates.**"` That is the one-object rule in plain English. The prompt is silent on titles and on `description`, as claimed | "the prompt does say *Never leave duplicates*, but never says the grader decides sameness by keyword substring over `title + description`" — which is **sharper**, because G-2 shows the rule fires on *distinct* obligations sharing vocabulary in 57 of 57 cases. The model was told the right rule and penalised by a different one |

**One observation no section made, offered for phase 3/4.** `runner.py:101` instructs the
model: *"Over-acting is the most common mistake; when in doubt, do nothing."* The system
prompt therefore actively steers every model toward the behaviour that collects V-3's
38.3% null floor — 56 no-action passes plus 8 free cancels. V-3 establishes the floor is
*available*; the prompt makes taking it the instructed strategy. That belongs in whatever
V-3 becomes.

---

## 4. Internal contradictions

1. **O-1's "can ever be re-graded" vs G-2, A-1 and A-6.** See 2c(i). The register performs
   three offline re-grades of paid runs and then says none is possible.
2. **The hazard box's "permanently destroys" vs the register's own git archaeology.** See
   2c(ii). C-6 and correction 19 both establish that everything under `corpus/` is tracked
   on every ref.
3. **`:1206` "K measured 5 and 21 only" vs K-2's table at `:920-927`,** which reports the
   defect sweep at `daily_max` 3, 5, 8, 13, 21 **and** 30. The proposed "free" experiment
   (`:1207`) is partly already done and its result is printed two pages earlier. What is
   genuinely un-run is the narrower free-typed-date-agreement test, which is what the
   sentence probably means; as written it is false.
4. **`:1235-1236` "proves no `get_email` was dropped" vs `:1267-1268` "neither confirmed it".**
   The body states as proved what the open-questions bullet lists as an unconfirmed
   assumption.
5. **C-5, count corrected without a note.** Register `:821` says "Four places record a
   truncated corpus sha"; `C.md:327` says "Three places" while its own evidence block
   (`C.md:336-339`) lists four. The register silently repaired a section-file error — right
   answer, undocumented delta.
6. **Two anchors silently corrected across sections.** `sb/oracle.py:51`→`:52` (G.md stale,
   register right) and `grader.py:163`→`:151` (K.md stale, register right). Both are
   verified right against the code; the register attributes them to the evidence brief
   (`:1399`, `:1402`) rather than to the section files, so a reader diffing G.md or K.md
   against the register finds an unexplained delta.

---

## 5. The "not reconciled" items — are they genuinely open?

| open question (`:1493-1508`) | verdict |
|---|---|
| **The isolated cwd and the answer key** | **Half settleable today, and I settled half of it.** The tool inventory of all four logs contains only the seven MCP tools plus `ToolSearch` — no `Read`, `Bash`, `Grep` or `Glob` on any day of any run (`grep '^   tools' <log> \| tr ',' '\n' \| sort \| uniq -c`). So "no log shows any model reading the corpus" is confirmed as far as the trace goes; it is *not* proof, because the trace is lossy for sonnet and haiku (O-2). The code half — whether the cwd blocks a built-in — is a static read of `runner.py:318-322` and needs no run. Genuinely open: only whether a model *could*, which one bounded smoke settles. Register's framing is fair. |
| **Whether the CLI emits one `assistant` event per content block** | **Genuinely open, and correctly identified as the cheapest live check.** But see 2a: the register spends this assumption in the body as though it were already settled. |
| **Whether `codex` behaves comparably** | **Genuinely open.** `codex` is not installed; `which codex` returns nothing. Not settleable here. |
| **`daily_max` 5 vs 21 for phase 5** | **Genuinely open — it is a decision, not a measurement.** But note item 1 in §3: the register's framing pre-loads it toward 21 with a claim ("only coherent at 21") its own table contradicts. |

---

## 6. Merge fidelity

Audited independently (full read of the register plus all six section files, mechanical ID /
severity / anchor / number diff).

- **All 50 IDs survive**, in all three places (section headings, Dashboard `:154-203`,
  per-finding subsections `:357-1150`). G 10, A 6, O 9, C 9, K 8, V 8. No renumbering, no
  invented IDs, no dropped IDs.
- **Severity: zero drift**, three-way (section `Severity:` line = Dashboard column =
  per-finding line) on all 50. Cost-to-verify likewise; A-5 is correctly the sole
  `needs-one-live-run`.
- **Counts line (`:205-207`) arithmetic is right**: 18 blocks / 26 distorts / 6 slows = 50;
  49 free-offline + 1 live = 50. I recounted from the register's own per-finding lines.
- **No finding lost its sole support.** 131 distinct `file:line` citations resolve. Three
  findings (A-1, C-6, K-4) carry no `file:line` at all, but their section files carry none
  either — nothing was dropped in the merge.
- **Number drift: one substantive** (item 6 in §3, K-7's "both" for "two of four"), plus the
  two undocumented anchor corrections and the C-5 count repair listed in §4.
- **Two lossy compressions worth knowing about:** G-4's table drops a seventh row present at
  `G.md:315` (`P2 + a realistic description | 138/167 | 82/111`); C's preamble narrows the
  "stamped nowhere" list from `--start/--days/--limit/--reasoning/--timeout` (`C.md:15-16`)
  to `--start/--days/--limit` (`:756`). Neither changes a retained number.

### Code anchors

**~145 `file:line` references opened at the cited lines. 138 CORRECT, 6 off-by-N, 0 WRONG.**
Nothing cited is fabricated; no reference points at a non-existent file or line. Every
load-bearing anchor in the fan-out checks out verbatim, including the ones the register
itself corrected against the brief: `grader.py:68-70`, `:151-152`, `:155-160`, `:164-165`,
`:168`+`:172-173` (**same `not title_set` guard confirmed — G-9's tautology claim is sound**),
`:186-195`; `runner.py:183`, `:141-143`, `:498-510`, `:513-516`, `:530`, `:535-558`;
`store_app.py:86-92` (**no third branch — A-4's blind spot confirmed**), `:139-144`,
`:192-213`, `:225-230`; `schema.py:81`, `:155`, `:393`, `:551-572`, `:563`; `oracle.py:52`.

Off-by-N (all point at the right *email or construct*, wrong line):

| cited | actual |
|---|---|
| `pizza-party.json:12` (party keyed `kind: todo`) | `:22-23` — `:12` is the email's `"id"` |
| `press-tour.json:110` (the leaked authoring note) | `:112` — the note is in the body |
| `Day-of-execution_and_Aftermath.json:36` (copied `dom:10,+2m`) | `:44` — `:36` is the `"id"` |
| `sb/resolver.py:16` ("serve-relative by design") | `:15` carries the sentence |
| `sb/scheduler.py:108` ("returning `False` forever") | `:109` is the `return False` |
| `sb/analyze.py:76-83` ("rebuilds the serve plan") | `:76-83` is the argparse block; the rebuild is `:88-90` (V-8 cites `:88-91` and is right) |

Three phase-0 anchors (`run.sh:18`, `run.sh:26`, `runner.py:499-514`) are stale at HEAD and
**correct against `24331fb`**, the revision they describe. That is appropriate for
"what was broken" entries but should be labelled as such.

**All 8 grep assertions re-run and confirmed:** `sid_filter` → 1 hit, the signature at
`runner.py:137`; `_node_state|_turn_delta|by_eid|eid_new` → 6 hits, all in `runner.py`, none
in tests; `total_cost_usd|duration_ms|num_turns|usage` → 2 hits, both in `_limit_reset_wait`
(one is the docstring, not the regex — a one-word imprecision); `\.tier\b` → **0**;
`sha256|hexdigest|blake2` → **0**; `find . -name "*.log"` → **0**; zero `tolerance` keys
anywhere in `corpus/` so all 134 ops take the `exact_day` default; `test_grader.py` has
exactly 4 tests and `grep '_title_hit\|count_ok\|_grade_op' sb/tests/` → **0 hits**.

Also confirmed: `24331fb` is 2026-07-26, message `.`, **26 files, +9257/−25** (C-6, exact);
`git log --all -- corpus/` goes `4f24e2a` → `24331fb`; `docs/_repair/*.md` total exactly
**4,368 lines**; exactly **one** attribution warning exists across all four logs
(`outputs/opus.md:67`), which is A-4/A-6's premise; exactly **two** bare `{serve}` tokens
corpus-wide (correction 13).

Two further small numeric errors:

- **C-9, `:876`:** "`:18-19` still reads 'SMOKE VERIFIED' **eleven lines** below the banner".
  The banner is `BENCHMARK_RESULTS.md:3-12`, so it is **six** lines below its end. Line refs
  right, distance wrong.
- **G-8's "21 of 134", settled.** A static count of `"match"` keys in `corpus/nodes/*.json`
  gives 114 present / **20** absent, which looks like a contradiction. It is not: one further
  op, `Sponsoring-Marathon.launching-sponsoring-eugene-marathon-2`, writes
  `"match": ["launchmeeting"]` where its obligation name *is* `launchmeeting`, so it takes the
  default in effect. 20 + 1 = **21**. The register's number is right under an
  "effective default" reading it never states; state it, or the number reads as an error.

---

## 7. Everything else I re-derived (all CONFIRMED)

Independently reproduced, exact unless noted:

- **G-2** — 57 duplicate failures (opus 14, sonnet 12, haiku 17, sonnet176 14) and **0 of 57**
  involve two objects sharing a title. `scratchpad/v_g2.py`.
- **G-8** — **21** ops take the effective default `match` (20 omit the key; one,
  `Sponsoring-Marathon.launching-sponsoring-eugene-marathon-2`, writes `["launchmeeting"]`
  which *is* its name — the register's 21 is right only under this "effective default"
  reading, which it does not state). Pass 8/63 = 13% vs 146/339 = 43%; length buckets
  35% / 45% / 41% / **3/42 = 7%**, all exact; 0 ops carry more than one keyword; 134/134 use
  the `exact_day` default.
- **G-6** — monotone decay confirmed (see §3 item 4 for the bucketing caveat).
- **G-3** — description-only matches 23/101, 14/97, 8/97 exact (sonnet176 2/85 by my parse vs
  2/91).
- **G-10** — `passed = count_ok and len(matched) >= 1` at `grader.py:165` does gate the date
  predicate behind the string match, as claimed.
- **A-3 / A-4 / A-5** — 46 of 56 no-action emails share a day with an acting email; 150/167
  emails have a same-day sibling in a different node; opus 51/52 and sonnet 46/47
  "(nothing matching created)" details on multi-email days; email ids mean 45 / max 77 chars,
  108 over 40 chars. Minor: "16 ids are a strict prefix of another" is **16 ordered pairs
  across 14 distinct prefix ids**.
- **C-3** — 93 feasible settings in C.md's grid, **12** score 166/167, all failing
  `Marketing-campaign-new-product-delay.serena-williams-reschedule`. Exact.
  `scratchpad/v_c3.py`.
- **C-9** — `daily_max=5` first feasible at `--days 57`, `=4` at `68`, and 1, 2, 3 infeasible
  at every `n_days` up to 2000. Exact.
- **K-1** — 143/167 emails carry no `date` edge; 49/167 (29%) served on a weekend. Exact.
  My conservative past-narrating-email detector (only `resolver.human()`-shaped dates) gives
  14/13/15/5/3 against K's 18/16/17/6/3 — a strict lower bound, same shape, and I cannot
  reproduce K's free-typed-date detector without its rule.
- **K-4** — Innovation-comp 48/167 emails, 17/134 ops, **zero T3**, 31/56 no-action; score
  share 36% / 40% / 37%; acting-only 35% / 36% / 38%; `Partnership-with-deeptech-companies`
  **2/10 for all three**; `World_Cup_Cleat_Launch` 7/8/8 of 22. All exact.
- **K-6** — 12 `by:` windows; mean **18.5**d, max **63**d (19.5 / 64 under inclusive day
  counting — a convention difference, not an error).
- **K-7** — 14 of 112 `eq` ops land on a Saturday or Sunday. Exact.
- **K-8** — 11 of 47 emitted anchors referenced by nothing; the `new_party_date` /
  `New_party_date` case pair; 13 distinct subjects appear more than once
  (`"Reveal event date and venue"` ×3); **10 op sites carry edge whitespace across 9 distinct
  names** (the register's "nine" is right at the name level). Minor: "3 emails have an empty
  `from`" — there are **4** (`Innovation-comp.a-kid-drew-the-new-logo-concept`,
  `Sponsoring-Marathon.pitch-deck-2`, `World_Cup_Cleat_Launch.wc-cleat-launch-window-options`,
  `pizza-party.pizza-place-selection`).
- **V-4** — exactly **190** `why` details in each 167-email log, 199 in the 176-email log.
- **V-5** — `--filler` 30, 60 and 200 raise `InfeasibleSchedule` at `--days 300` while 90,
  120, 150 and 175 succeed; 175 has a *lower* mean needle span (59.5) than 150 (62.6). The
  error's advice is wrong exactly as described — it says "raise `--days` … try >= 40" when
  `--days 300` was supplied.
- **V-6 / V-7** — tiers 50 T1 / 67 T2 / 50 T3; T3 needles closer than T2 on both axes
  (email-span 31.1 vs 33.3, day-span 10.4 vs 11.8); 32 of 50 T3 emails have no answer-key
  anchor reference; 13 of 50 are pure no-action. All exact. **Caveat the register omits:** the
  span comparison rests on **n=6 T2 needles against n=18 T3**. V-7's conclusion should carry
  that n.
- **O-6** — narration snippets sitting at the 200-char cap on **57/57, 50/57, 16/16, 18/19**
  days. Exact.
- **Corrections 16, 19, 20** — opus↔sonnet verdict agreement **148/167 = 88.6%**; em-dash
  share 61% / 0% / 0%; mean title length 47 / 36 / 40; unanimous passes 77, unanimous fails
  58, discriminating **32/167 (19%)**; op level 95/134 = 71% at 0/3 or 3/3, 62 ops failed by
  every model; action emails 39 / 40 / 42 with 58 unanimous failures. Exactly four files in
  the entire history contain a `SCORE n/m` line (`git grep -l -E "SCORE [0-9]+/[0-9]+"
  $(git rev-list --all)`), none of them 51%; `git stash list` empty; `git fsck --lost-found`
  reports nothing.

---

## 8. What I could NOT check, and why

1. **Anything requiring a live model.** A-5 (the only `needs-one-live-run` finding), M-5's
   inherited-environment question, and the raw-`stream-json` capture that would settle
   whether the CLI emits one `assistant` event per content block. Prohibited by the brief and
   by `CLAUDE.md`.
2. **The `codex` half of the roster.** `codex` is not installed on this machine, so
   `_parse_codex`'s resolved-model path (C-7), its cost fields (O-8) and its non-deduping
   trace (O-2) are unexercised. Confirms the register's own open question rather than
   resolving it.
3. **K-3's eight prose-contradicts-answer emails and K-5's rendering defects.** Both rest on
   hand-reading prose against resolved dates. I reproduced the *mechanical* half (the
   `Day-of-execution_and_Aftermath` copied-expression case is checkable and real) but not the
   hand-classification, and K states its own false-positive rate honestly (9 flagged, 4 true).
   I did not re-do the hand pass.
4. **K-1's 18 past-narrating emails and K-2's "prose contradicts its answer" column.** My
   detector only sees `resolver.human()`-shaped dates and returns a strict lower bound
   (14 at `daily_max=5`). Reproducing 18 needs K's free-typed-date rule, which is not in the
   register.
5. **C-5's six failed hash reconstructions.** I did not attempt a seventh algorithm. The
   register bounds this claim honestly already ("a seventh algorithm … could still match").
6. **C-6's commit-content claims** beyond the ones I spot-checked, and **C-8's overwritten
   84/176 haiku run**, which by construction leaves no artifact.
7. **Whether the two destructive scripts actually run to completion**, since running them is
   forbidden. I read them; I did not execute them. My git-recoverability finding rests on the
   tracked/clean state of `corpus/`, not on observing a destructive run.
8. **The webapp, the production DB and the `backups` branch** — off limits per `CLAUDE.md`;
   `origin/backups` was read only through `git grep`/`git rev-list`, never checked out.

---

## 9. Recommended edits, in priority order

1. **Downgrade phase 1 from "time-critical".** `corpus/` is git-tracked and clean; the
   recovery is not perishable. Re-argue the phase order on G-4 alone (which holds).
2. **Restate O-1.** "No paid run can be re-graded under a rule that would newly admit objects
   the log never rendered" — and cite G-2/A-1/A-6 as the counter-examples that define the
   boundary.
3. **Fix C-1's haiku levers** to `urgency_horizon ≤ 7` (seven values, not three) and restore
   C.md's grid definition next to the number 785.
4. **Soften resolution 2's "proves"** to "meets the lower bound on 56 of 57 days", drop the
   serialisation claim, and change correction #8 to "a tight lower bound for opus".
5. **Fix K-7's "both"** → "two of the node's four graded emails".
6. **Retitle K-2** away from "only coherent at `daily_max=21`" — its own table shows 30 is
   better.
7. **Drop G-3's writing-style causal clause**; it is unmeasurable from any surviving artifact
   by G-3's and O-5's own arguments.
8. **Fix G-1's prompt claim** — `runner.py:107` does say "Never leave duplicates". The
   sharper version (told the right rule, graded by a different one) strengthens G-1 and
   connects it to G-2.
9. **Add the ceiling table from §2b** to Contradiction 1, replacing the correlational
   sentence.
10. **Add the bucketing rule to G-6**, the eq/any_of restriction to G-2's counterfactual, and
    the "effective default" reading to G-8's 21, so each number travels with its method.
11. **Note the n=6 vs n=18** behind V-7's span comparison.
12. **Housekeeping:** six off-by-N anchors (§6); C-9's "eleven lines" → six; K-8's "3 emails
    have an empty `from`" → **4**; A-4's "16 ids are a strict prefix of another" → 16 *pairs*
    across 14 ids; label the three phase-0 anchors as `@24331fb`.

**Nothing in this pass changes a severity, and nothing found here would move a finding out of
`open`.** The register's evidentiary base is sound; what needs editing is the prose that
over-reads it, and the phase table that was built on the one claim that does not hold.

---

## 10. One thing outside the register's scope, found while checking it

**The entire phase-A output is uncommitted, and the register says so about itself in C-8.**

```
git status --porcelain          ->  M docs/benchmark-repair.md
                                    ?? docs/_repair/
git show 67b3005:docs/benchmark-repair.md | wc -l   ->  239
wc -l docs/benchmark-repair.md                      ->  1590
git check-ignore -v docs/_repair/                   ->  exit 1 (NOT ignored, just never added)
```

The committed register is the 239-line pre-merge version. The 1,590-line merged register and
all **4,368 lines** of `docs/_repair/{G,A,O,C,K,V}.md` — the working the register explicitly
delegates to ("**Go to the section file for the working**", `:66-67`) — exist only in the
working tree. This is C-8's exact mechanism ("artifacts preserved only by hand") operating on
the fan-out's own evidence, and it is a larger exposure than the corpus hazard the register
puts in its hazard box: `corpus/` is committed and recoverable, `docs/_repair/` is not
recoverable from anything. **Commit `docs/_repair/` and the merged register before doing
anything else.**
