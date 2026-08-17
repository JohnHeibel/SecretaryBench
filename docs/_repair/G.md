# G — grader identity

Scope: object matching (`_title_hit`), the kind filter, description in the match
haystack, cumulative-pool collisions, lint check #5's blind spot, and the oracle's
inability to detect an ungradeable answer key.

All measurements below are offline and free. Four recorded runs were parsed
(`outputs/opus.md`, `outputs/sonnet.md`, `past/claude-haiku-4-5.md`, and
`past/claude-sonnet-4-5.md` — see the "brief corrections" section: the fourth is not in
the brief's §2.1 table). Parsed email counts and totals match each log's own `SCORE`
line exactly: 90/167, 91/167, 98/167, 102/176.

---

## G-1 the match keyword is often not derivable from any text the model reads
Status: open
Severity: blocks-measurement
Cost to verify: free-offline

**What's wrong.** `_title_hit` (`sb/grader.py:68-70`) requires every keyword in `op.match`
to appear as a lowercased substring of `f"{obj.title} {obj.description}"`. For 32 of the
125 graded create/move ops (26%) the required keyword appears **nowhere** in the rendered
mail of the entire node, so the model has no evidence the string is wanted. The benchmark
therefore scores string-guessing, not scheduling: pooled over the three 167-corpus runs,
obligations whose keyword is present in the mail pass 48% of the time, and obligations
whose keyword is absent pass 8%.

**Evidence.**
- Grader contract: `sb/grader.py:68-70`.
- Keyword derivability (renders every body with the real serve context, then substring-tests
  each keyword against `sender + subject + rendered body` for the whole node):

```
.venv/bin/python - <<'EOF'
from datetime import date
from sb import resolver
from sb.resolver import Context
from sb.schema import load_corpus
from sb.scheduler import build_plan
c = load_corpus("corpus"); p = build_plan(c, start_date=date(2026,6,1), seed=42, n_days=60)
txt = {}
for n, nd in c.nodes.items():
    txt[n] = " ".join(f"{e.sender} {e.subject} "
                      f"{resolver.render_body(e.body, Context(p.serve_date[e.id], p.anchors)).text}"
                      for e in nd.emails).lower()
ops = [(e, op) for e in c.emails.values() for op in e.answer.ops if op.verb != "cancel"]
bad = [(e.node, op.name, op.match) for e, op in ops
       if not all(k.lower().strip() in txt[e.node] for k in op.match)]
print(len(bad), "of", len(ops), "graded create/move ops have an undiscoverable keyword")
for b in sorted(bad): print("  ", b)
EOF
```
  → `32 of 125`. Examples: `OpenAI In Person Walk Through` wants `['through']`;
  `Pitch Breifing` wants `['breifing']` (an authoring typo — the mail says "briefing");
  `LeBron James marketing campaign scheduled` wants that whole phrase;
  `Giano Ronaldo marketing campaign ` wants that phrase including a trailing space;
  `approve_prize_amounts` wants `['approve']`, absent from the node.
- A second class is created by rendering: `{!name = expr}` tokens render to a date and the
  anchor *name* is discarded (`sb/resolver.py:345-356`). An author reading the JSON sees
  `delayed`, `signoff`, `conference` "in the body"; the model never does. Verified: those
  three strings occur in `corpus/nodes/*.json` but zero times in any rendered body/subject
  of their node.
- Effect size, pooled over `opus` + `sonnet` + `haiku`, at the level of individual graded
  ops (join each `expected  <kind> ~"<kw>" @ …` log line back to its corpus op):

  | | keyword present in the node's mail | keyword absent |
  |---|---|---|
  | opus | 43/101 = 43% | 4/33 = 12% |
  | sonnet | 51/101 = 50% | 2/33 = 6% |
  | haiku | 52/101 = 51% | 2/33 = 6% |
  | **pooled** | **146/303 = 48%** | **8/99 = 8%** |

  The authored keyword predicts the outcome (48% vs 8%) far more strongly than the model
  does (43% vs 50% vs 51%).
- 62 of the 134 ops were failed by all three models; 28 of those 62 (45%) have an
  undiscoverable keyword.
- Worked example, `outputs/opus.md:926-935` — one email, three ops, three failures, no
  visible model error:

```
 ✗ FAIL  [138] Enterprise_Ai_Selection.ai-meeting-schedule  · served Fri Jul 17
   ✗ expected  event ~"anthropic" @ Tue Jul 21
     actual    "Anthropic vendor session (Zoom)" Tue Jul 21 10 AM (60m); "Google Gemini tour (Zoom)" Tue Jul 21 2 PM (60m)
     why       found 2 matching, expected exactly 1 (duplicate / double-booked)
   ✗ expected  event ~"google" @ Tue Jul 21
     actual    "Anthropic vendor session (Zoom)" Tue Jul 21 10 AM (60m); "Google Gemini tour (Zoom)" Tue Jul 21 2 PM (60m)
     why       found 2 matching, expected exactly 1 (duplicate / double-booked)
   ✗ expected  event ~"through" @ Wed Jul 22
     actual    (nothing matching created)
     why       no event titled like "through" was created
```

- The model is never told the rule. `SYSTEM_PROMPT` (`sb/live/runner.py:87-107`) says
  nothing about titles being keyword-matched, nothing about the description field, and
  nothing about one-object-per-obligation.

**Why it matters for benchmark validity.** A benchmark whose stated contribution is
retrieval span is, on a quarter of its graded obligations, measuring whether the model
independently guessed a private string. The score is not a lower bound on capability
either, because the same mechanism produces false passes (see G-3).

**Options.**
1. **Publish the contract to the model.** Extend the system prompt / MCP tool docs so
   titles are told to name the obligation, or add an explicit `obligation` argument to
   `create_event`/`create_todo`. Tradeoff: makes the task partly about following a naming
   convention, changes the tool surface the model sees (the same objection that blocked
   the mcp 2.0 port, register E-2), and invalidates comparison with the recorded runs.
2. **Author-side constraint.** Add a lint that rejects any `match` keyword absent from the
   node's *rendered* mail. Tradeoff: free and mechanical, but it only fixes discoverability
   — it does nothing about the collisions in G-2, and it makes an existing 32-op corpus
   edit mandatory before the next run (a K-category cost).
3. **Replace substring matching with semantic matching** (embedding similarity, or an
   LLM judge asked "is this object the obligation described here?"). Tradeoff: removes the
   string-guessing axis entirely and handles paraphrase; costs money per grade, introduces
   a nondeterministic grader, and needs its own validation against a hand-graded baseline.
4. **Grade by attribution instead of by title** — require the model to reference the
   obligation it is satisfying (e.g. one object per `email_id` per kind), and drop title
   matching. Tradeoff: eliminates the whole problem class but collapses multi-obligation
   emails, and moves the failure mode into A's `email_id` routing.

**Overlaps with:** G-3, G-4, G-5, G-8; K (the 32 keywords are an authoring artifact);
A (attribution as an alternative identity channel).

**Open questions.** Of the 8 passes on undiscoverable keywords, how many were coincidence
(e.g. a model naturally writing "Approve …") versus a keyword that *is* discoverable via a
route this measurement misses (an ancestor node, the sender name)? Does any keyword become
discoverable only through `search_inbox`, which would make it a legitimate retrieval test?

---

## G-2 the exactly-one rule counts keyword hits, not obligations
Status: open
Severity: blocks-measurement
Cost to verify: free-offline

**What's wrong.** `sb/grader.py:164-165` sets `count_ok = len(title_set) == 1` where
`title_set` is every object in the node's cumulative same-kind pool whose text contains the
keyword. Because obligations inside one storyline share vocabulary, a *correct* second
object routinely lands in a *different* obligation's `title_set` and fails it. Across all
four recorded runs, **not one** of the 57 `found N matching, expected exactly 1
(duplicate / double-booked)` failures involves two objects with the same title — every one
is a collision between distinct obligations, which is not what the message says.

**Evidence.**
- `sb/grader.py:151-152, 164-165, 175`.
- Same-title check over all four logs (parse each `actual` line into `"<title>" <when>`
  records, split on `; (?=")` to survive quotes inside titles):
  `opus 0/14, sonnet 0/12, haiku 0/17, sonnet176 0/14` duplicate-failures contain a
  repeated title.
- The colliding sets are the storyline's own vocabulary, and they repeat across models —
  `atlas`, `design`, `launch`, `reveal`, `sponsor`, `board`, `outsole`, `trophy`, `list`,
  `marketing`. Quoted from `outputs/opus.md`:
  - `expected  event ~"reveal" @ …` / `why  found 5 matching` — the five are
    `"World Cup cleat reveal event — CEO on stage"`, `"WC cleat press briefing (under
    embargo)"`, `"WC cleat reveal rehearsal — run-through with the striker"`,
    `"Manufacturing kickoff — WC cleat (factory hears timeline from you)"`,
    `"Design Lead 1:1 — credit for the design team on the boot"`.
  - `expected  event ~"atlas"` in `project_atlas`, whose every object is about Atlas.
- The model was often demonstrably right. Comparing each collide failure's object dates
  against the expected day parsed out of the `expected` line (restricted to `eq`/`any_of`
  predicates, where the day is unambiguous): opus 9/14, sonnet 7/10, haiku 12/16,
  sonnet176 10/13 collide failures contained an object **on an expected day**. At the
  email level, the emails whose *only* failing detail is such a case:

  | run | emails | score today | score if the rule were "at least one on the right day" |
  |---|---|---|---|
  | opus | 7 | 90/167 (54%) | 97/167 (58%) |
  | sonnet | 7 | 91/167 (54%) | 98/167 (59%) |
  | haiku | 11 | 98/167 (59%) | 109/167 (65%) |
  | sonnet176 | 9 | 102/176 (58%) | 111/176 (63%) |

  Note the ordering flips nothing but the gap widens — this is not a monotone rescale.
- Even a scheduling-perfect agent trips it. Running the shipped engine with an oracle that
  is correct in every respect but titles objects with the author's own obligation name,
  7 of its 10 failures are collide (see G-4's table, policy P2): e.g.
  `expected event ~"reveal" @ Mon Jul 20 / actual "Reveal Event" Mon Jul 20 9 AM;
  "Reveal Rehearsal" Sun Jul 19 9 AM`.

**Why it matters for benchmark validity.** The rule exists to catch a real behaviour —
a reschedule that leaves a stale duplicate (`sb/grader.py:11-13`). It is currently firing
almost entirely on a different phenomenon, and the log labels the result "duplicate /
double-booked", which is the opposite of what happened. Any reader of these logs will
conclude models leave duplicates. Measured: they did not, in 57/57 cases.

**Options.**
1. **Best-match assignment.** Give each obligation the single best-scoring object and
   forbid one object from satisfying two obligations (Hungarian / greedy by score).
   Tradeoff: makes grading order-independent and kills the collision class; is a real
   algorithm change needing its own tests, and makes "the model left a duplicate" harder
   to express.
2. **Keep exactly-one but scope it to the turn**, not the run: require exactly one match
   among objects created/updated for this obligation, and check the cumulative pool only
   for the stale-duplicate condition after a `move`. Tradeoff: much smaller change; still
   collides whenever two obligations are created in the same node on nearby days.
3. **Relax to "≥1 matching object satisfies the predicate, and no *stale* object of this
   obligation survives"**, where staleness is tracked by object identity across a `move`.
   Tradeoff: needs object identity, which the state-based design deliberately avoids
   (`sb/grader.py:4-6`), and would not survive delete-then-recreate.
4. **Leave the rule and fix the corpus** so no two same-node same-kind obligations can
   share vocabulary. Tradeoff: no grader change (so a score move is attributable), but
   G-1's measurement shows models write the storyline's shared nouns regardless of what
   the author picks; this pushes authors toward artificial keywords, which is exactly what
   produced the 32 undiscoverable ones.

**Overlaps with:** G-1, G-3, G-5, G-6, G-7; K (vocabulary overlap is authored).

**Open questions.** How many of the 57 collide failures would also fail under best-match
assignment because the model genuinely created two objects for one obligation? That needs
the object-level dump from phase 1 (O), not the log.

---

## G-3 `description` is in the match haystack, invisible in the log, and unmentioned to the model
Status: open
Severity: blocks-measurement
Cost to verify: free-offline

**What's wrong.** `sb/grader.py:69` builds the haystack as
`f"{obj.title} {obj.description}".lower()`. The MCP tools expose `description` on
`create_event` and `create_todo` (`sb/live/mcp_app.py:42,73`), the system prompt never
mentions it (`sb/live/runner.py:87-107`), and `_fmt_obj` (`sb/grader.py:125-129`) prints
only the title — so descriptions silently decide matches and no log shows it. In
`outputs/opus.md`, 23 of the 101 objects the grader attributed to an obligation (23%) have
a title that does not contain the required keyword; they matched through the description
alone.

**Evidence.**
- `sb/grader.py:69` and `sb/grader.py:125-129`; `sb/live/mcp_app.py:42,51,73,82`.
- Description-only matches, counted by testing each logged object's *title* against the
  op's keywords (a title-only miss on an object the grader accepted implies the
  description carried the keyword):

  | run | objects attributed | title lacks ≥1 keyword | share |
  |---|---|---|---|
  | opus | 101 | 23 | 23% |
  | sonnet | 97 | 14 | 14% |
  | haiku | 97 | 8 | 8% |
  | sonnet176 | 91 | 2 | 2% |

  The spread tracks how verbose each model's descriptions are, which means the metric is
  partly measuring writing style.
- These cause **false failures**: 7 of opus's 14 collide failures collapse to exactly one
  object if the description is dropped from the haystack (sonnet 1/12, haiku 4/17,
  sonnet176 0/14). The clearest is `outputs/opus.md:926-932`, quoted in G-1 — the Anthropic
  event and the Google event each caught the other's keyword through its description.
- They also cause **false passes**. `outputs/opus.md`:
  `expected to-do ~"location" @ … / actual "Contact retreat venue to approve the plan" /
  why matched` — no "location" in the title.
  `past/claude-haiku-4-5.md`: `~"livestream"` matched
  `"Live launch stream - walk through new ACME"`. Two such passes in opus, two in sonnet,
  three in haiku.
- Controlled experiment, offline, using the shipped `sb.engine` + `sb.grader`: take the
  **unmodified oracle** (`" ".join(op.match)` as the title, i.e. a guaranteed keyword hit)
  and add a realistic description (`subject + first 180 chars of the rendered body`) to
  every object it creates. Score falls from **167/167 (100%) to 141/167 (84%)**, and 25 of
  the 29 lost details are `collide`. Nothing about the scheduling changed.

**Why it matters for benchmark validity.** The one field the grader reads that the model
is never told about is a field models fill with the email text — which is, by construction,
the storyline's shared vocabulary. It converts every verbose description into a
cross-obligation collision, and it is invisible in every artifact we have.

**Options.**
1. **Match on `title` only.** Tradeoff: one-line change, removes 23% of opus's matched
   objects from consideration in both directions (some of them were passes); cheap to
   evaluate offline once phase 1 (O) lands a state dump; changes the score, so it must not
   share a commit with any corpus edit.
2. **Keep description but weight it** — require the keyword in the title, allow the
   description only to disambiguate ties. Tradeoff: more forgiving of paraphrase; makes the
   contract two-tier and harder to state in one sentence.
3. **Keep the haystack and tell the model** what it is used for. Tradeoff: no grader change
   (score moves remain attributable), but it turns the description into a gaming surface —
   the cheapest strategy becomes stuffing every plausible keyword into every description,
   which `docs/POTENTIAL_GAMING.md` should then cover.
4. **Print the description in `_fmt_obj`.** Diagnostic only, does not change any score.
   Tradeoff: costs nothing and makes the other three decisions measurable from a log
   instead of by inference; log volume grows.

**Overlaps with:** G-1, G-2, G-9; O (the log omits the field that decided the match).

**Open questions.** Was `description` intended to be part of identity at all, or is
`f"{title} {description}"` a leftover from a fuzzier matching idea? Nothing in
`docs/ANSWER_KEY_GRAMMAR.md` or `sb/schema.py:53-63`'s `Op` docstring ("title keywords")
suggests descriptions were meant to count.

---

## G-4 the oracle certifies satisfiability and is structurally blind to gradeability
Status: open
Severity: blocks-measurement
Cost to verify: free-offline

**What's wrong.** `sb/oracle.py:51` titles every object `" ".join(op.match)`, so the
reference model writes the answer key's own keywords into the calendar and scores 100% on
any corpus, however unmatchable its keywords are. `sb/scale.py:126-127` prints that 100% as
"must be 100% — corpus is valid at scale", which is the check standing between the corpus
and a paid run. Replacing only the title policy — leaving the scheduling perfect — drops the
same oracle to as low as 55%, the same band as the three real models.

**Evidence.**
- `sb/oracle.py:51,55,60`; `sb/scale.py:126-127`; `sb/engine.py:133-150` (the loop reused).
- Title-policy sweep, run through the shipped `sb.engine.run` + `sb.grader` on `corpus/`
  with `seed=42, start=2026-06-01, n_days=60` (identical plan to the recorded runs; the
  agent is scheduling-perfect in every variant — right kind, right day, right verb, right
  node, no stale duplicates):

  | title policy | score | action-emails only (n=111) |
  |---|---|---|
  | P0 `" ".join(op.match)` — shipped oracle | 167/167 **100%** | 111/111 100% |
  | P1 `op.name` verbatim | 160/167 96% | 104/111 94% |
  | P2 `op.name` humanized (`_`/`-` → space) | 157/167 94% | 101/111 91% |
  | P3 the email's subject line | 92/167 **55%** | 36/111 32% |
  | P4 subject + humanized name | 140/167 84% | 84/111 76% |
  | P0 + a realistic description | 141/167 84% | 85/111 77% |
  | P2 + a realistic description | 138/167 83% | 82/111 74% |

  Recorded models for comparison: opus 54%, sonnet 54%, haiku 59%.
- P3 is not a strawman: titling a calendar entry after the email's subject is a normal
  assistant behaviour, and the system prompt gives no reason to do otherwise. A
  **scheduling-perfect** agent scores 55% — inside the 54–59% band of the three real runs.
  This does not prove the models were perfect; it proves the observed band is fully
  reachable without a single scheduling error.
- Even the author's own name for the obligation costs 6–10 points (P1/P2), so the corpus is
  not gradeable by an agent that knows exactly what each obligation *is* but not what
  string the grader wants.
- The oracle's `find_in_node` (`sb/engine.py:69-74`) searches with the same joined keyword
  string it wrote, so an oracle `cancel` deletes exactly what it created and can never hit
  the cross-obligation collision that fails real models (G-7).

**Why it matters for benchmark validity.** "The corpus lints clean and every answer key is
satisfiable" (register, corpus health check) is true and does not mean what it is being used
to mean. The gate that is supposed to catch a bad answer key is the one component
guaranteed to pass it.

**Options.**
1. **Add a second gate: a paraphrase oracle.** Re-run the corpus with P2 (and/or P3) and
   fail CI below a threshold. Tradeoff: free, deterministic, catches G-1/G-2/G-5 at
   authoring time; the threshold is arbitrary and P2 is only a proxy for what models write.
2. **Score the corpus against titles real models actually produced.** The four logs give
   ~270 distinct observed titles keyed by node; replay them through the grader.
   Tradeoff: the most realistic vocabulary available, and free; the sample is biased toward
   objects that already matched something, and it cannot cover obligations no model ever
   attempted.
3. **Keep one oracle but randomize the title within the contract** — e.g. keyword plus
   filler, or keyword embedded in a longer natural phrase. Tradeoff: minimal change,
   catches nothing about discoverability; it only tests that the substring rule tolerates
   surrounding text, which it already does.
4. **Leave the oracle alone and rename its output** so `sb.scale`'s line reads
   "answer keys are satisfiable" rather than "corpus is valid". Tradeoff: costs nothing and
   removes a false assurance, but adds no new signal.

**Overlaps with:** G-1, G-2, G-5, G-7; V (this is the mechanism that let construct
problems ship); C (`sb.scale`'s 100% is quoted as a validity claim in the register).

**Open questions.** P3's 55% and the models' 54–59% could be coincidence; the right test is
whether a P2/P3 oracle and a real model fail the *same* obligations. That comparison is
possible offline today against the recorded logs and was not run here.

---

## G-5 lint check #5 compares author strings to author strings
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** `sb/schema.py:551-572` flags a collision only when every keyword of one
obligation's `match` set is a substring of some keyword of a sibling's `match` set. Both
sides are author-invented; the grader matches against titles the *model* invents. The
corpus lints clean, yet a variant of the same check that compares each `match` set against
its siblings' own obligation **names** — the author's own description of the object, and the
closest free proxy for a model's title — flags 10 collisions.

**Evidence.**
- `sb/schema.py:559-572`, and its own comment at `:556-558`: "Authors invent match keywords;
  the model invents titles, so this can only be prevented here, at authoring time." The
  check does not act on that observation — it never looks at anything a model would write.
- Shipped check: 0 flags. Name-aware variant: 10 flags, in 7 of the 15 nodes:

```
Company-Retreat       ['list']      (Create VIP List)        also catches 'Create A List For Athlete Visit'
Enterprise_Ai_Sel…    ['ai']        (AI_Final_Meeting_Review) also catches 'AI_Sign_Off'
Innovation-comp       ['board']     (board_deck_one_pager)   also catches 'review_board_slides'
Partnership-with…     ['fbs']       (FBS Planning Meeting)   also catches 'FBS Conference'
Pre-Launch            ['sign']      (sign_COO_docs)          also catches 'production_handoff_design'
World_Cup_Cleat…      ['delivery']  (Outsole sample delivery) also catches 'Confirm Outsole PO'
World_Cup_Cleat…      ['reveal']    (Reveal Event)           also catches 'Reveal Rehearsal'
project_atlas         ['freeze']    (Atlas code freeze)      also catches 'Atlas board demo'
project_atlas         ['interview'] (Atlas press interview)  also catches 'Atlas board demo'
project_atlas         ['atlas']     (Atlas board demo)       also catches 'Atlas launch dinner'
```
  `['atlas']` alone catches three of its four sibling names. Every one of these appears as
  a real `found N matching` failure in the logs (G-2).
- A second blind spot: keywords that are substrings of ordinary words in their own node's
  rendered mail. Measured by tokenising each node's rendered bodies and subjects:
  `'ai'` is inside `email` and `openai` (its own node contains an `OpenAI` obligation);
  `'end'` inside `attend`, `attendance`, `send`; `'sign'` inside `design` and `sign-off`
  (its own node contains `production_handoff_design`); `'launch'` inside `launch-night`,
  `launch-week`, `launches`; `'design'` inside `designer`, `designers`;
  `'william'` inside `williams`; `'first'`, `'floor'`, `'thank'`, `'teaser'`, `'spad'`,
  `'pitch'`, `'sponsor'`, `'green'` likewise. 23 such cases across the corpus.
- The check is also restricted to `verb == "create"` (`sb/schema.py:563`), so a `cancel`
  keyword is never examined even though cancel is the verb most sensitive to collisions
  (G-7).

**Why it matters for benchmark validity.** The linter is the only automated defence between
an author and an ungradeable answer key, and it is testing a relation that cannot detect
the failure it was written to prevent. The corpus's clean lint is currently cited as
evidence of corpus health.

**Options.**
1. **Widen the probe set** to `{match keywords} ∪ {obligation name} ∪ {significant words of
   the subject}` for each sibling, same node and kind, and flag any containment.
   Tradeoff: free and catches the 10 above; will also fire on cases that are fine in
   practice, so it needs a suppression mechanism or it will block authoring.
2. **Lint against observed model vocabulary.** Keep a checked-in file of titles harvested
   from recorded runs and flag any `match` set that catches more than one object in its
   node. Tradeoff: grounded in reality rather than proxy; the vocabulary file becomes an
   artifact to maintain, and it is empty for a new storyline.
3. **Make the check a warning with a count** rather than a hard `CorpusError`.
   Tradeoff: unblocks authoring and surfaces risk; a warning nobody reads is worth nothing,
   and the register's own rule is that nothing moves on vibes.
4. **Delete check #5 and rely on whatever replaces the exactly-one rule** (G-2 option 1).
   Tradeoff: no false alarms; loses the only authoring-time signal if the grader change
   turns out not to eliminate collisions.

**Overlaps with:** G-1, G-2, G-7, G-8; K (this is the guardrail K's authoring drift walked
past).

**Open questions.** Should the check be per-kind (as now) or global? The pool is kind-
filtered (`sb/grader.py:151`), so cross-kind collisions are harmless for matching, but
`athlete` is currently the `match` for both `Athlete Meeting` (event) and
`Create A List For Athlete Visit` (todo) in one node — safe today, and instantly broken by
any change that unifies the pools.

---

## G-6 the pool is cumulative over the whole run, so obligations get harder by position
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** `_grade_op` matches against `state.events` / `state.todos`
(`sb/grader.py:151`), which the live runner builds from the *entire* store filtered only by
node (`sb/live/runner.py:137-146`). Every object created anywhere in a storyline stays in
the pool forever, so the probability of a collision rises monotonically with how late an
obligation is served. Pooled over the three 167-corpus runs, pass rate falls from 52% to
26% as the pool fills.

**Evidence.**
- `sb/grader.py:151`; `sb/live/runner.py:137-146`.
- `_node_state` takes a fourth argument `sid_filter: set[str]` and never references it
  (`sb/live/runner.py:137-146`); the call site passes `eid_new`
  (`sb/live/runner.py:509`). The cumulative behaviour is intended per the module docstring
  (`sb/grader.py:9-13`), but the dead parameter means the call site reads as if the pool
  were scoped to the turn. Confirmed dead: `grep -rn "sid_filter" sb/` returns only the
  definition line.
- Pass rate by pool depth, counting how many same-node same-kind `create` obligations were
  already due when the op was graded (serve order from
  `build_plan(corpus, start_date=2026-06-01, seed=42, n_days=60)`):

  | prior same-kind obligations in the pool | pass rate (3 runs pooled) |
  |---|---|
  | 0 | 42/81 = 52% |
  | 1–2 | 65/156 = 42% |
  | 3–5 | 33/111 = 30% |
  | 6+ | 14/54 = 26% |

- Objects whose `email_id` does not resolve to a corpus email are dropped from **every**
  node's pool (`sb/live/runner.py:140-142`, `corpus.emails.get(o.email_id)`), so a mistyped
  or invented id makes an object invisible to the grader and produces
  `(nothing matching created)`.

**Why it matters for benchmark validity.** A metric that decays with position inside a
storyline confounds difficulty with position. `Innovation-comp` is 48 emails (register
§2.10 / brief §2.10), so its late obligations are graded against the deepest pool in the
corpus — the node with the most emails is also the node where the grader is least reliable.

**Options.**
1. **Keep cumulative.** It is what makes "a reschedule that leaves a stale duplicate fails"
   expressible (`sb/grader.py:11-13`). Tradeoff: no change and no risk; the decay stays and
   must be reported as a caveat on every score.
2. **Window the pool** to objects created or last touched within N days of the serve date.
   Tradeoff: bounds the decay; N is arbitrary and a genuinely stale duplicate created long
   ago would escape.
3. **Scope the pool to the obligation's own lifetime** — from the create's serve date to the
   op being graded. Tradeoff: principled, and removes collisions with obligations that did
   not exist yet; needs per-obligation bookkeeping the state-based design does not carry.
4. **Leave the pool and fix the count rule instead** (G-2). Tradeoff: one change addresses
   both; if the count rule becomes best-match, pool depth stops mattering — but that is an
   assumption, not a measurement.

**Overlaps with:** G-2, G-7; K (`Innovation-comp` dominance); A (invalid `email_id`
handling); O (no state dump means this can only be inferred).

**Open questions.** The 52%→26% decay has a confound: obligations served later in a
storyline may genuinely be harder (more anchors, more retrieval). Separating position from
difficulty needs the tier tag, which `sb/analyze.py` never reads (brief §2.9) — a V/O
dependency.

---

## G-7 `cancel` is graded as keyword absence over the cumulative pool
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** `sb/grader.py:155-160` passes a `cancel` only when **zero** objects in the
node's cumulative same-kind pool contain the keyword. Any sibling object the model was
supposed to *keep* — or any over-action elsewhere in the node — makes the cancel
unpassable. Every cancel failure in all four recorded runs is of this form: the object that
"survived" is a different obligation.

**Evidence.**
- `sb/grader.py:156-158`.
- All 13 cancel failures across the four runs, quoted from the logs:
  - `~"launch"` (cancel `Atlas launch dinner`) survived as
    `"Project Atlas — code freeze go/no-go"` and `"Project Atlas — public launch"` (opus),
    `"Project Atlas go/no-go (code freeze)"` (sonnet),
    `"First press interview - post-launch"` (haiku, sonnet176).
  - `~"design"` (cancel `Design cut meeting`) survived as
    `"Private meeting with Design Lead - team structure discussion"` (haiku),
    `"Design team future discussion with Melissa"` (sonnet).
  - `~"sponsor"` (cancel `sponsor_call`) survived as
    `"Sponsor mixer with retail partners (optional appearance)"` (opus) — an object created
    for a **no-action** email, so one over-action cascaded into a second failure.
  - `~"dynamics"` (cancel `Boston Dynamics Visit`) survived as
    `"Boston tech trip (WHOOP + Boston Dynamics)"` in all four runs — the models modelled
    the trip as one event, which the corpus splits into five obligations.
- Pass rates: opus 6/9, sonnet 5/9, haiku 5/9, sonnet176 6/9.
- The oracle cannot reproduce this: `sb/oracle.py:54-56` deletes by the same joined-keyword
  string it wrote, so its cancels are always clean (G-4).
- `sb/schema.py:563` restricts lint #5 to `verb == "create"`, so no authoring-time check
  ever examines a cancel keyword against its siblings.

**Why it matters for benchmark validity.** Cancel is 9 of 134 ops (7%) and is the verb where
the grader's rule is strictest. A third to a half of cancel failures are the grader
objecting to objects the answer key requires to exist.

**Options.**
1. **Grade cancel against the same assignment used for create/move** (G-2 option 1): the
   cancel passes if the object *assigned* to that obligation is gone. Tradeoff: consistent
   with the rest of the grader; depends entirely on G-2 landing first.
2. **Require absence only among objects previously attributed to that obligation** — track
   which object satisfied the create. Tradeoff: precise; needs identity across turns, which
   the state-based design avoids.
3. **Keep absence but scope it to objects on the cancelled obligation's expected date.**
   Tradeoff: cheap; wrong whenever the model cancelled by moving, and does nothing for the
   `dynamics` case.
4. **Leave it and treat the Boston-trip class as a corpus problem** — five obligations for
   one trip is an authoring choice models consistently reject. Tradeoff: no grader change;
   does not touch the `launch`/`design`/`sponsor` collisions, which are pure keyword
   overlap.

**Overlaps with:** G-2, G-5, G-6; A (the `sponsor` case starts as an over-action on a
no-action email); K (trip granularity).

**Open questions.** In the `dynamics` case, is one merged trip event the *right* answer? If
so this is a K finding about obligation granularity, not a G finding, and the grader is
correctly reporting a real disagreement.

---

## G-8 `match` defaults to the whole obligation name as one contiguous phrase
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** `sb/schema.py:154-156` sets `match=list(raw.get("match", [])) or [name]`,
so an author who omits `match` gets the entire obligation name — spaces, underscores,
hyphens, punctuation and all — as a single required substring. 21 of 134 ops (16%) do this,
and they pass at 13% versus 43% for explicit keywords. Nothing lints the `match` field at
all.

**Evidence.**
- `sb/schema.py:154-156`; the `Op` docstring at `sb/schema.py:53-56` documents the default.
- Defaults in the corpus include `Team_pizza_party`, `order-the-pizzas`, `Design Lead 1:1`,
  `sponsorship & budget approval meeting`, `sponsorshippitch`, `launchmeeting`,
  `Retreat Company Meeting Call`, `Approve revised event budget`,
  `LeBron James marketing campaign scheduled`, and
  `Giano Ronaldo marketing campaign ` (with a trailing space — the only whitespace-dirty
  keyword; `_title_hit` does not strip, `sb/grader.py:70`, and it works today only because
  `f"{title} {description}"` always appends a space).
- Pooled over the three 167-corpus runs, at op level:
  default `[name]` **8/63 = 13%**; explicit `match` **146/339 = 43%**.
- By keyword length, same pooling: ≤5 chars 26/75 = 35%; 6–9 chars 99/222 = 45%;
  10–15 chars 26/63 = 41%; ≥16 chars (phrases) **3/42 = 7%**.
- All 134 ops have exactly one keyword
  (`Counter(len(op.match) …) == {1: 134}`), so the multi-keyword conjunction the schema
  offers for disambiguation is entirely unused — which is the natural fix for G-2 and no
  author has reached for it.
- `sb/schema.py:130-156` (`_parse_op`) validates verb, kind, predicate and tolerance, and
  performs no validation of `match`: no non-empty check, no whitespace strip, no minimum
  length. A `match: [""]` would match every object (`"" in hay` is always true) and pass
  the linter.

**Why it matters for benchmark validity.** The path of least authoring effort — omit
`match` — produces the least gradeable keyword in the corpus, and the metric silently
records that as model failure. The 7% pass rate on ≥16-char keywords is the strongest
single predictor of failure found in this category apart from G-1's discoverability split.

**Options.**
1. **Remove the default; make `match` required** on every `create`. Tradeoff: forces the
   author to make the decision consciously; breaks every existing node file until edited
   (21 ops) and is a schema change with a webapp editor consequence
   (`sb/schema.py:30-39` feeds `schema.generated.ts`).
2. **Keep the default but tokenise it** — split the name on whitespace/underscores/hyphens
   into a keyword list, so `Team_pizza_party` becomes `["team","pizza","party"]`.
   Tradeoff: a one-line change that makes the default behave like a reasonable author;
   turns some defaults into over-broad filters (`["design","lead","1:1"]` is narrower,
   `["approve","revised","event","budget"]` much broader).
3. **Lint the `match` field**: reject empty/whitespace-only keywords, strip whitespace,
   warn above a length threshold and below ~3 characters. Tradeoff: free and mechanical;
   thresholds are arbitrary and it does not help the 21 ops that need rewriting.
4. **Leave it and treat the 21 ops as a K-category corpus edit.** Tradeoff: no code change,
   so a subsequent score move is attributable to the corpus alone (the register's
   never-both-in-one-commit rule); the trap stays armed for the next author.

**Overlaps with:** G-1 (16 of the 20 non-cancel default-match ops are also undiscoverable),
G-5; K.

**Open questions.** Was the default ever intended for production authoring, or only as a
convenience for tests and fixtures? `sb/tests/fixtures` would show which.

---

## G-9 the grader's own explanation cannot distinguish its failure modes
Status: open
Severity: slows-work
Cost to verify: free-offline

**What's wrong.** `(nothing matching created)` (`sb/grader.py:168`) and
`no <kind> titled like "X" was created` (`sb/grader.py:173`) fire under the *same*
condition (`not title_set`), so the brief's observation that all 52 opus cases show
`(nothing matching created)` is a tautology of the code, not evidence. `_fmt_obj`
(`sb/grader.py:125-129`) prints neither the description that decided the match, nor the
object's kind, nor the `email_id` it was attributed to, so no recorded log can attribute a
failure to a cause.

**Evidence.**
- `sb/grader.py:168` and `:172-173` share the guard `elif not title_set` / `if title_set
  else`. The two strings can never disagree.
- `sb/grader.py:125-129`: `_fmt_obj` emits title + when (+ duration for events) only.
- The distinguishable causes, all reachable, none separable from the log:
  (a) genuine under-action; (b) a title/description lacking the keyword; (c) a kind
  mismatch (`sb/grader.py:151` filters the pool by `op.kind`); (d) an `email_id` naming an
  email in another node (`sb/live/runner.py:140-142`); (e) an `email_id` that resolves to
  nothing, which drops the object from every pool; (f) action taken on a later day than the
  email was graded — grading is one-shot at `sb/live/runner.py:503-511` and never revisited.
- Best available offline split of the 52/47/52 cases (see the "what §4.3 now looks like"
  section below): 52–60% are on ops whose keyword appears nowhere in the node's mail;
  25–31% unexplained; 13–15% have a keyword-matching object under another node; 0–4% are a
  proven kind mismatch. Cause (f) measures at 1/52 in each run and appears to be a
  coincidence match on `['ai']`, so late action is not a significant contributor.

**Why it matters for benchmark validity.** Every fix in this category changes the score, and
without a cause-tagged artifact there is no way to check whether it moved the score toward
the truth. This is the concrete reason the register sequences O and phase 1.5 before G.

**Options.**
1. **Emit a machine-readable per-op record** — `{obligation, kind, keywords, pool_size,
   title_set, matched, cause}` — alongside the human log. Tradeoff: makes every later
   grader change re-scorable offline; is squarely O's work, not G's, and must not change
   any verdict in the same commit.
2. **Split the reason strings by cause** inside `_grade_op`: distinguish "pool empty" from
   "pool non-empty, no keyword hit", and tag whether the hit came from the title or the
   description. Tradeoff: small, local, immediately useful; only reaches causes (a)/(b)/(c),
   since (d)/(e)/(f) are invisible to the grader by construction.
3. **Print the description and `email_id` in `_fmt_obj`.** Tradeoff: cheapest possible
   change and settles G-3 from logs rather than by inference; makes logs substantially
   longer and changes the format `sb/analyze.py:25` re-parses.
4. **Do nothing in G and rely entirely on phase 1's state dump.** Tradeoff: avoids touching
   grader code before the honesty baseline exists; leaves the four logs already paid for
   permanently un-attributable.

**Overlaps with:** G-1, G-3, G-6; O (machine-readable output); A (causes d/e).

**Open questions.** Is any recorded run's store recoverable? `sb/live/runner.py:513-516`
terminates the store in `finally` with no dump, so probably not — which would mean cause
attribution for the existing four logs is permanently capped at what is measured above.

---

## G-10 the identity logic has no unit tests
Status: open
Severity: slows-work
Cost to verify: free-offline

**What's wrong.** `sb/tests/test_grader.py` contains four tests, all about date predicates
and the oracle's blackout handling. Nothing tests `_title_hit`, the description in the
haystack, the kind filter, the exactly-one rule, or the cancel rule — the code responsible
for roughly 85% of recorded failures (brief §2.3, reproduced exactly below).

**Evidence.**
- `grep -n "def test" sb/tests/test_grader.py` →
  `test_by_predicate_rejects_due_date_before_email_arrives`,
  `test_by_predicate_accepts_due_date_between_serve_and_deadline`,
  `test_oracle_target_avoids_not_in_blackout`,
  `test_oracle_target_errors_when_window_fully_blocked`.
- The date machinery those four cover decides only about half the outcomes:
  `passed = count_ok and len(matched) >= 1` (`sb/grader.py:165`) reaches the predicate only
  when `len(title_set) == 1`, i.e. exactly the `matched` + `on the wrong day` details.
  Of 125 create/move details per run: opus 41+18 = 59 (47%), sonnet 48+18 = 66 (53%),
  haiku 49+7 = 56 (45%). **On 47–55% of graded obligations the answer key's date is never
  evaluated** — the outcome is settled by string matching before any temporal reasoning is
  tested.
- `op.tolerance` is `exact_day` on 134/134 ops, and predicates are 112 `eq`, 12 `by`,
  1 `any_of`, 9 none — so the tested surface is also the least varied part of the corpus.

**Why it matters for benchmark validity.** Every option in G-1…G-8 changes `_grade_op`.
With no test pinning current behaviour, "the score went up" and "the grader broke" are
indistinguishable, and the register's rule that nothing reaches `verified` without a named
artifact has nothing to name.

**Options.**
1. **Characterisation tests first** — pin today's behaviour (collision fails, description
   counts, cancel-by-absence) before changing anything, so a later diff shows exactly which
   verdicts moved. Tradeoff: cheap and makes every later change attributable; codifies
   behaviour we may be about to delete.
2. **Property tests** over synthetic pools: "one object matching one obligation always
   passes", "adding an unrelated object never changes a verdict" — the second currently
   fails, which is the bug. Tradeoff: states the intended contract rather than the current
   one; requires deciding the contract first, which is G-2's open decision.
3. **Golden-file regression** on a fixture corpus + fixture store, asserting the exact
   `details` list. Tradeoff: highest coverage per line written; brittle against formatting
   changes, and formatting is what G-9 wants to change.

**Overlaps with:** all of G; O.

**Open questions.** Should the characterisation tests be written against the current
formatting (which G-9 proposes to change) or against the structured record O will
introduce? Writing them twice is waste; writing them against a format that does not exist
yet blocks on O.

---

## What §4.3 now looks like (the split of the "nothing matching created" cases)

Each failing detail line was joined back to its corpus op by
`(email_id, kind, "/".join(op.match))`, then classified against three offline signals:
S2 — an object satisfying the keywords is visible in the log under the same node but the
*other* kind (proven kind mismatch); S1 — the keyword appears nowhere in the node's
rendered mail (no correct behaviour reliably produces it); S3 — a satisfying object is
visible under a *different* node.

| | opus (52) | sonnet (47) | haiku (52) |
|---|---|---|---|
| S1 keyword absent from the node's mail | 27 (52%) | 28 (60%) | 29 (56%) |
| unexplained (under-action or paraphrase) | 16 (31%) | 13 (28%) | 13 (25%) |
| S3 same-keyword object under another node | 8 (15%) | 6 (13%) | 8 (15%) |
| S2 proven kind mismatch | 1 (2%) | 0 | 2 (4%) |

Every case mapped; 0 unmappable in all three runs.

Read this as an **upper bound on model fault**: S1 does not prove the model created
something wrongly titled (it may have created nothing), but it does prove that in 52–60% of
these failures the grader demanded a string the model had no way to know. Genuine
under-action is therefore at most 31%/28%/25% of the bucket, not the whole of it. S3 is the
weakest signal — a shared keyword such as `atlas` or `design` can legitimately appear in
another storyline's objects — so it should be treated as "worth checking", not proven.
The three proven kind mismatches are `Innovation-comp.let-s-set-up-a-recap`
(`recap_meeting` keyed `event`), `Enterprise_Ai_Selection.final-review`
(`AI_Final_Meeting_Review` keyed `todo`), and `Company-Retreat.athelete-visit`
(`Contact People Added To List` keyed `event`) — so the kind filter, which the brief lists
as one of four equal candidates, is the smallest contributor by an order of magnitude.

---

## Brief corrections and confirmations

**Reproduced exactly.** §2.3's whole failure-mode table, for all three runs, detail for
detail (52/47/52 not-found, 14/12/17 duplicate, 18/18/7 wrong-day, 5/5/0 over-acted,
3/4/4 cancel-left, 51/51/56 no-action passes, 41/48/49 matched, 6/5/5 cancelled).
§2.2's forensics (em-dash share 61%/0%/0% against the brief's 64%/0%/0% — a small
methodology difference, same conclusion; mean title length 47/36/40 against 46/36/40;
opus↔sonnet verdict agreement 148/167 = 88.6% exactly). §2.4's grader mechanics and every
`file:line` in §3 that this category touches. §2.10's oracle 100%, reproduced independently
through `sb.engine.run` rather than `sb.scale`.

**§2.1 is incomplete — there is a fourth recorded run.**
`past/claude-sonnet-4-5.md` (1047 lines) is a complete log: header
`Corpus: 176 emails (sha 809d389794dd79a9) · seed 42 · days 30 · daily_max 21`,
`Score: SCORE 102/176 (58%)`, `Tally: PASS 102 · FAIL 74 · ERROR 0 · search_inbox 1`,
176 per-email verdicts, all parseable. The brief says `BENCHMARK_RESULTS.md` §5's "two
evidence links point at files that do not exist" — the *links* (`outputs/claude-*.md`) are
indeed dead, but the sonnet half of that evidence survives under `past/`. It is a second
corpus and it shows the same signature (14 collide failures, 0 of them same-title; the same
`atlas` / `design` / `launch` / `dynamics` / `ai` collisions), which is what makes the
signature corpus-independent rather than an artifact of the current 167-email corpus.

**§2.3's "in all 52 opus cases the `actual` field reads `(nothing matching created)`" is a
tautology, not an observation.** `sb/grader.py:168` writes that string under exactly the
condition (`not title_set`) that `sb/grader.py:172-173` uses to write the corresponding
`why`. It could not have come out any other way, and it is not evidence of anything. The
sentence that follows it — that the log cannot distinguish the causes — is correct.

**§4.5 is now established, with a stronger statistic than verdict agreement.** Across the
three 167-corpus runs: 77/167 emails (46%) are unanimous passes, 58/167 (35%) unanimous
fails, and only **32/167 (19%) discriminate between the models at all**. At op level,
95/134 (71%) are 0/3 or 3/3, and 62 ops are failed by every model. Restricting to the 111
action emails, the three models score 39/111, 40/111, 42/111 (35%/36%/38%) and 58 of those
111 are unanimous failures. The benchmark's entire discriminating range is 32 emails; the
observed 8-point spread sits inside a band that is 81% pre-determined. §4.1 is right that
verdict agreement cannot discriminate the two hypotheses — this measurement replaces it.

**§4.1's "roughly 100 of the 167 outcomes are determined by the grader" is directionally
right but the arithmetic should be restated.** 135/167 outcomes are *model-invariant*,
which is not the same as *grader-determined* — three models could genuinely all fail a hard
obligation. The defensible causal subset is: the 57 collide failures (0/57 involving a
repeated title), the ≥7-per-run collide failures where a correct object sat on the correct
day, the 99 pooled op-details whose keyword is undiscoverable and which pass at 8%, and the
25 description-driven collisions the shipped oracle suffers at 100% scheduling accuracy.

**§2.5 is now quantified.** The shipped oracle scores 100%; the same oracle, scheduling
unchanged, scores 96% titling by the obligation name, 94% humanized, 84% with a realistic
description added, and 55% titling by the email subject. A scheduling-perfect agent lands
inside the 54–59% band of the three real models.

**§2.4's `grader.py:163-165` anchor for the pool** should be `grader.py:151` (`pool = …`);
`:163-165` is `matched` / `count_ok` / `passed`. Substantively the description is right.

**Not challenged, but noted:** §2.12's claim that no model read the answer key is consistent
with everything measured here — every collision and every miss is explainable by the
grader's own mechanics, and a model that had read `corpus/nodes/*.json` would have written
`breifing` and `Team_pizza_party` into its titles. None did.
