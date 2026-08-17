# A — attribution and no-action grading

Scope: how a day's created objects are routed back to the email that should own them
via the model-supplied `email_id`; the escapes from the no-action check; the
monitor-only `invalid_email_id` / `stale_email_id` warnings.

All measurements below are offline, on the committed logs and corpus, at the levers the
`outputs/` runs used (`seed 42`, `start 2026-06-01`, `days 60`, `daily_max 5`, which
reproduces the 57-day / 167-email plan those logs show).

**Where the model's `email_id` actually reaches the score.** Only two channels:

1. `runner.py:141-143` — `_node_state` keeps an object only if
   `corpus.emails.get(obj.email_id).node == node`. The stamp therefore selects which
   node pool the object lands in, and an unknown id drops it from every pool.
2. `runner.py:507-510` — `eid_new` (objects created today *and* stamped with this
   email's id) is passed as `TurnDelta`, and `grader.py:186-195` uses that delta
   **only** in the no-action branch.

For the 111 emails that carry ops, the stamp matters solely as a node selector; the
`TurnDelta` is never read (`grader.py:197` ignores `turn`). For the 56 no-action emails
the stamp is the whole grade. The third argument `sid_filter` at `runner.py:137` is
never referenced in the function body (`grep -rn "sid_filter" sb/` returns exactly one
hit, the signature), so the per-email id set has no effect on op grading at all.

---

## A-1 More than half of every recorded score is a no-action verdict, and the model ranking is not robust to the attribution rule
Status: open
Severity: blocks-measurement
Cost to verify: free-offline

**What's wrong.** 51/90, 51/91 and 56/98 of the passes in the three recorded runs are
"correctly took no action", i.e. 56-57% of every headline score is the abstain check
rather than the scheduling check. That check's only input is the model's self-reported
`email_id` plus same-day timing, and re-grading the same logs under a day-scoped
attribution rule instead of an id-scoped one inverts the model ranking. The benchmark
therefore cannot currently claim to rank models, because the ranking is a property of
the attribution rule and not of the models.

**Evidence.**
- Per-email split of each score (parsed from the committed logs; totals match the
  headline scores exactly):

  | log | total | no-action passes | acting passes | no-action share of passes |
  |---|---|---|---|---|
  | `outputs/opus.md` | 90/167 | 51/56 | 39/111 | 57% |
  | `outputs/sonnet.md` | 91/167 | 51/56 | 40/111 | 56% |
  | `past/claude-haiku-4-5.md` | 98/167 | 56/56 | 42/111 | 57% |

  ```
  for f in outputs/opus.md outputs/sonnet.md past/claude-haiku-4-5.md; do
    echo "$f pass=$(grep -c '✓ PASS' $f) noaction_pass=$(grep -c 'why *  *correctly took no action' $f) overacted=$(grep -c 'why *  *over-acted' $f)"
  done
  # outputs/opus.md pass=90 noaction_pass=51 overacted=5
  # outputs/sonnet.md pass=91 noaction_pass=51 overacted=5
  # past/claude-haiku-4-5.md pass=98 noaction_pass=56 overacted=0
  ```
- Counterfactual, day-scoped rule (any object created during the day counts against
  every no-action email in that day's batch, instead of only objects stamped with that
  email's id). Number of no-action passes that sit on a day where the model demonstrably
  called `create_event`/`create_todo`:

  ```
  for f in outputs/opus.md outputs/sonnet.md past/claude-haiku-4-5.md; do
    printf "%-30s " "$f"
    awk '/── day /{c=0} /^ *tools /{c=(/create_event|create_todo/)?1:0} /correctly took no action/{n++; if(c)k++} END{printf "no-action PASS=%d, on a day with a visible create_*=%d\n", n, k}' "$f"
  done
  # outputs/opus.md              no-action PASS=51, on a day with a visible create_*=35
  # outputs/sonnet.md            no-action PASS=51, on a day with a visible create_*=34
  # past/claude-haiku-4-5.md     no-action PASS=56, on a day with a visible create_*=55
  ```
  Applying that: opus 90→55 (33%), sonnet 91→57 (34%), haiku 98→43 (26%). The order
  goes from `haiku 98 > sonnet 91 > opus 90` to `sonnet 57 > opus 55 > haiku 43`.
  Because the tool trace is lossy (`runner.py:180`, brief §2.7), the `create_*` counts
  are lower bounds, so the drops are lower bounds too.
- The central anomaly of §2.1 is majority no-action: haiku beats opus by 8 points, of
  which **5 are no-action** (56/56 vs 51/56) and 3 are acting emails (42/111 vs 39/111).
- 56 of 167 emails carry no ops, and **31 of those 56 (55%) are in `Innovation-comp`**:
  ```
  PYTHONPATH=$PWD .venv/bin/python -c "
  from sb.schema import load_corpus
  c=load_corpus('corpus'); na=[e for e,m in c.emails.items() if not m.answer.ops]
  print(len(na), sum(1 for e in na if c.emails[e].node=='Innovation-comp'))"
  # 56 31
  ```

**Why it matters for benchmark validity.** The stated contribution is retrieval span
over a temporal-reasoning task. The majority of the reported score is instead an
abstention check concentrated in one storyline, graded by a rule the ranking is not
robust to. Any headline number in `BENCHMARK_RESULTS.md` is therefore a blend of two
different measurements with unequal weights that nobody chose.

**Options.**
1. Report no-action and acting accuracy as two separate numbers and stop publishing a
   single blended score. Cheap and free-offline, changes no grading; but it makes the
   headline number disappear, and every existing external reference to "54%" becomes
   unanchored.
2. Keep one score but reweight so the two classes contribute in a chosen ratio.
   Preserves a single figure and lets the ratio be a stated design decision; introduces
   a tuning knob that will be read as score inflation, and the reweighted numbers are
   not comparable to any published run.
3. Rebalance the corpus so no-action is a stated fraction (currently 33.5% of emails,
   57% of passes) and is spread across nodes rather than 55% in one. Fixes the cause
   rather than the report; it is a corpus change, so it invalidates comparison with all
   three recorded runs and cannot be done in the same commit as a grader change.
4. Leave the weighting and fix only the attribution rule (A-3), then re-measure whether
   the ranking is still rule-dependent. Smallest change; but the 57% concentration and
   the single-node concentration both survive.

**Overlaps with:** G (the acting half of the score is dominated by keyword matching),
V (binary per-email metric, tier data unread), K (`Innovation-comp` dominance),
C (levers not stamped, so the haiku `daily_max=21` confound is invisible in the log).

**Open questions.** Is the day-scoped rule in the counterfactual even the right
alternative? It is deliberately extreme (a day containing acting emails *should*
produce objects), so it measures sensitivity, not correctness. Does `daily_max` itself
move no-action accuracy? Within opus+sonnet at `daily_max=5` the no-action pass rate by
batch size is 6/8 (batch 1), 16/18, 26/26, 31/34, 23/26; haiku at `daily_max=21` scored
56/56 across batches of 9-21. Suggestive of a positive relationship, but n=8 at batch 1
makes it far from established.

---

## A-2 The live attribution split is untested and semantically different from the offline path the oracle validates
Status: open
Severity: blocks-measurement
Cost to verify: free-offline

**What's wrong.** Attribution is implemented twice with different semantics: `sb/engine.py`
grades one email per turn with an unfiltered delta, while `sb/live/runner.py` grades one
*day* per turn and splits the delta by model-supplied `email_id`. Only the offline one is
exercised by tests and by the oracle, so `oracle: 167/167 = 100%` is evidence about a code
path that no live run uses. Nothing in `sb/tests/` touches `_node_state`, `_turn_delta` or
`by_eid`.

**Evidence.**
- Offline: `engine.py:143-147` — `before = store.snapshot_ids()` / `model(...)` /
  `turn = store.delta_since(before)` inside `for email_id in batch`. The delta is
  per-email and unfiltered; `email_id` is never used to route it.
- Live: `runner.py:501-510` — `day_new = _all_ids(state) - before`,
  `by_eid = {r["id"]: r.get("email_id", "") ...}`,
  `eid_new = {i for i in day_new if by_eid.get(i) == eid}`.
- Divergent failure mode for an unknown id: `engine.py:97`
  `node = self._corpus.emails[email_id].node` raises `KeyError`; `runner.py:141-143`
  `e = corpus.emails.get(o.email_id)` returns `None` and the object is silently dropped.
  Same input, hard crash offline vs silent mis-score live.
- Dead parameter: `_node_state(corpus, state, node, sid_filter)` at `runner.py:137`;
  `grep -rn "sid_filter" sb/` returns one line, the signature. `runner.py:509` passes
  `eid_new` into it and it is discarded.
- No coverage: `grep -rn "_node_state\|_turn_delta\|by_eid\|eid_new" sb/` returns only
  `sb/live/runner.py`. `sb/tests/test_e2e.py:9` imports `from sb.engine import Store, run`,
  so all three e2e tests run the offline path. `sb/tests/test_live_store.py:6-38` covers
  only that the store *emits* the two warnings.
- Consequence for the corpus health check in the register ("oracle: 167/167 = 100%"):
  `sb/oracle.py` feeds `engine.run`, which has no `email_id` routing, so a 100% oracle
  cannot detect any attribution defect.

**Why it matters for benchmark validity.** Every "the pipeline is clean" claim in the
register and in `BENCHMARK_RESULTS.md` rests on the oracle and the 62 unit tests, and
neither touches the code that decides which email gets credit for which object. The
live attribution rule has never been executed against a known-answer input.

**Options.**
1. Delete `sid_filter` and add unit tests for `_node_state` / `_turn_delta` /
   `eid_new` against synthetic store snapshots. Free-offline and non-behavioural; it
   pins current behaviour without deciding whether current behaviour is right, so it
   could entrench the very rule A-3 questions.
2. Collapse the two paths: have the live runner build `NodeState`/`TurnDelta` through
   `sb.engine.Store` (or the reverse) so one implementation serves both. Removes the
   divergence class entirely; it is a refactor across the offline/live boundary and
   would change what the oracle validates, so it needs its own before/after artifact.
3. Make the offline path day-based too, so the oracle exercises attribution. The oracle
   would then be able to fail on an ungradeable-by-attribution corpus; but the oracle
   titles objects from the answer key (`oracle.py:51`) and stamps the correct id by
   construction, so it would still not exercise a *wrong* stamp unless a deliberately
   mis-stamping mock model is added.
4. Leave both paths and document the divergence, treating the oracle as a corpus
   satisfiability check only. Zero cost; the register then has to stop citing oracle
   100% as evidence about the harness.

**Overlaps with:** O (no state dump means the live path's output cannot be replayed
against the offline path), G (`_grade_op`'s node pool is the other consumer of the
stamp).

**Open questions.** Was the per-day live split a deliberate divergence from the
per-email offline model, or drift? The runner docstring at `runner.py:11-12` and
`grader.py:16` ("Per-email until the day-loop lands; ADR 0001 stage note") suggest
the day loop was a planned migration that only landed on one side. `docs/adr/` should
be checked for what ADR 0001 actually committed to.

---

## A-3 The no-action check reads a day-scoped and id-scoped delta, giving five ways for over-action to escape
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** `grade_email` fails a no-action email only if `turn` is non-empty, and
`turn` contains exactly the objects that were created during *this day's single turn* and
stamped with *this email's exact id*. Everything else the model does for that email is
invisible, and the verdict is computed once, on the serve day, and never revisited. The
brief's §2.6 names one of the five escapes (the sibling stamp); the other four are
unnamed and one of them is the escape actually observed in the recorded runs.

**Evidence.** Sources for each route:
1. **Sibling stamp, same day.** `runner.py:507` `eid_new = {i for i in day_new if
   by_eid.get(i) == eid}`. An object stamped with any other id in today's batch is not
   in this email's delta. 46 of the 56 no-action emails are served alongside at least
   one acting email.
2. **Later day, any stamp.** `runner.py:501` `day_new = _all_ids(state) - before`, where
   `before` is read at `runner.py:431`, and grading runs inside the same day iteration
   (`runner.py:503-512`). An object created on a later day is never in any earlier day's
   delta, and the earlier email's verdict is already recorded in `results[eid]`.
   **This is the route observed in `outputs/opus.md`** (see A-6).
3. **Unknown or non-current stamp.** `store_app.py:89-92` warns but does not block;
   `runner.py:507` still excludes the object from the no-action delta.
4. **Create-then-delete inside the same turn.** `day_new` is a set difference over
   end-of-turn state (`runner.py:500-501`), not a log of create calls, so an object
   created and deleted within one turn leaves no trace.
5. **Patch an existing object instead of creating one.** `store_app.py:139-144`
   `patch_event` mutates in place, keeps the id, and calls no attribution check, so the
   object is not in `day_new` at all. A model can retitle and redate a stale object to
   serve a no-action email with zero grading consequence.
- Exposure measurement:
  ```
  PYTHONPATH=$PWD .venv/bin/python -c "
  from datetime import date
  from sb.schema import load_corpus
  from sb.scheduler import Levers, build_plan
  c=load_corpus('corpus'); p=build_plan(c,start_date=date(2026,6,1),seed=42,n_days=60,levers=Levers(1,5,7))
  d=[b for b in p.per_day if b]; na={e for e,m in c.emails.items() if not m.answer.ops}
  print('no-action with an acting same-day sibling:', sum(1 for b in d for e in b if e in na and any(o not in na for o in b if o!=e)))
  print('emails with a same-day sibling in a DIFFERENT node:', sum(1 for b in d for e in b if any(c.emails[o].node!=c.emails[e].node for o in b if o!=e)))"
  # no-action with an acting same-day sibling: 46
  # emails with a same-day sibling in a DIFFERENT node: 150
  ```
- Counter-evidence, stated for the verifier: all 10 observed over-action failures
  (5 opus, 5 sonnet) carry titles topically derived from the no-action email itself,
  e.g. `outputs/opus.md:445-448` `Sponsoring-Marathon.approval-needed-race-sponsorship-budget`
  → `"Get CFO to send VP Marketing the marathon master budget"`. Where the model over-acts
  it stamps correctly. There is no log evidence of an adversarial stamp in any run.

**Why it matters for benchmark validity.** The no-action check is 57% of the score
(A-1) and is described in `grader.py:15` as "the turn must create nothing". It in fact
tests "created nothing today under this exact id", which is a much weaker property, and
four of the five gaps require no wrong id at all. A model that batches its work across
days, or tidies up after itself, or edits instead of creating, scores better on abstention
than one that acts immediately and honestly.

**Options.**
1. Re-evaluate no-action at end of run rather than on the serve day: fail an email if any
   object stamped with its id exists in final state. Closes routes 2 and 3 and needs no
   store change; it cannot close routes 1, 4 and 5, and it delays every no-action verdict
   to the end, so a run that dies mid-way grades nothing.
2. Grade a day's objects as one pool against the day's batch (no per-email split): the
   day passes its no-action emails only if the day's created objects are all accounted
   for by the batch's acting emails. Closes routes 1 and 2 and removes the dependency on
   the model's stamp; it makes the metric day-granular rather than email-granular, which
   changes what a "167-email score" means, and by A-1's measurement it moves 34-55
   verdicts per run.
3. Grade from a create/update/delete *action log* instead of end-of-day state diff.
   Closes routes 4 and 5, and is the only option that can distinguish transient from
   persistent action; it requires the store to record an ordered log, and it partly
   reverses the deliberate state-based design in `grader.py:4-6`.
4. Harden at the tool boundary: reject `create_*` whose `email_id` is not in today's
   batch (the `docs/POTENTIAL_GAMING.md:59` "potential future fix"). Removes route 3
   and narrows route 2; it changes the tool contract the model is measured against, is
   invisible to routes 1, 4 and 5, and by A-6's arithmetic would have changed exactly one
   verdict across all three recorded runs.

**Overlaps with:** A-4 (only route 3 is monitored), A-6 (worked instance of route 2),
G (the escaped object still enters the cumulative node pool), O (routes 4 and 5 are
undetectable without an action log).

**Open questions.** Routes 4 and 5 have zero observed instances, but the artifacts cannot
show them: the tool trace is lossy (`runner.py:180`) and no state or action log is
persisted. Are they worth closing before there is any evidence they occur? Also unresolved:
what *should* the correct answer be for an email whose work the model legitimately defers
to the next day? The corpus has no vocabulary for "acceptable within N days".

---

## A-4 `_watch_attribution` is blind to the exact case it was built to detect, so "fired once" is not evidence
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** `_watch_attribution` warns only when the stamped id was never served
(`invalid_email_id`) or was served on an earlier day (`stale_email_id`); an id belonging
to a *different email in today's batch* falls through the `elif` and produces no warning
at all. That is precisely the sibling stamp that §2.6 cites the warning count as evidence
about. The brief's inference ("those warnings fired once, so this is a theoretical hole")
is therefore invalid: the monitor cannot observe the hole it is being read as evidence for.

**Evidence.**
- `store_app.py:86-92`:
  ```
  if email_id not in _served_email_ids():
      _warn("invalid_email_id", ...)
  elif email_id not in _today_email_ids():
      _warn("stale_email_id", ...)
  ```
  There is no third branch. A sibling id is in `_today_email_ids()`, so both conditions
  are false.
- Exposure: 150 of 167 emails have a same-day sibling in a *different node*, which is the
  damaging case (the object leaves the node pool it was meant for). See the command in A-3.
- Warning count across all committed logs is 1, and it is the only line:
  ```
  grep -n "invalid_email_id\|stale_email_id" outputs/*.md past/*.md
  # outputs/opus.md:67:   warning  stale_email_id (object_kind=event, email_id=Innovation-comp.sponsor-mixer-before-the-final, title=Sponsor mixer with retail partners (optional appearance))
  ```
- Asymmetric validation at the store: `store_app.py:225-230` `get_email` raises
  `HTTPException(404, ...)` for an unknown id, while `store_app.py:126-131` and
  `:156-161` accept any string as `email_id` on create. The model gets an error when it
  *reads* a wrong id and silence when it *writes* one.
- No aggregation anywhere. `runner.py:489-490` prints the day's new warnings and nothing
  else consumes them; the store is terminated at `runner.py:515` with no dump; and
  `sb/analyze.py` never reads a warning line (`grep -n "warning" sb/analyze.py` returns
  nothing). The end-of-run summary at `runner.py:525-532` has no warning count.
- Transcription hazard the monitor would catch but nothing measures: email ids average
  45 characters (max 77), 108 of 167 exceed 40 characters, and 16 ids are a strict prefix
  of another id. The model must copy one exactly per create call
  (`mcp_app.py:42-44`, `SYSTEM_PROMPT` at `runner.py:103`).
  ```
  PYTHONPATH=$PWD .venv/bin/python -c "
  from sb.schema import load_corpus
  i=list(load_corpus('corpus').emails); L=[len(x) for x in i]
  print('n=%d mean=%.1f max=%d over40=%d' % (len(i), sum(L)/len(L), max(L), sum(1 for x in L if x>40)))
  print('strict-prefix pairs:', sum(1 for a in i for b in i if a!=b and b.startswith(a)))"
  # n=167 mean=45.0 max=77 over40=108
  # strict-prefix pairs: 16
  ```

**Why it matters for benchmark validity.** `docs/POTENTIAL_GAMING.md:47-49` sets an
explicit decision rule: "monitor first in the pilot log, then harden if it appears". The
monitor as built cannot make that decision, because the damaging same-day case is silent
by construction, and the warnings it does emit are printed and discarded. Reading "one
warning in three runs" as reassurance is the failure mode the doc was written to avoid.

**Options.**
1. Add a `sibling_email_id` warning when the stamp is in today's batch but is not the id
   whose turn produced the object. Cheap; but the runner cannot tell which email a turn
   was "on" (one turn covers the whole day), so this may only be expressible as a
   node-mismatch heuristic, which would fire on legitimate cross-node work.
2. Warn when a created object's stamped id resolves to a node with no ops matching the
   object (a "no obligation this could serve" signal). Catches the damaging cross-node
   case without needing per-email turns; it leaks answer-key structure into the store,
   which currently knows nothing about the corpus.
3. Persist warnings into the run artifact and surface a count in the end-of-run summary
   and in `sb.analyze`. Free, non-behavioural, and makes the monitor-first rule
   executable; it does not widen coverage, so the blind spot stays.
4. Make the store validate `email_id` on write the way it already does on read (404 on
   unknown, and optionally on non-current). Removes the class rather than observing it;
   it changes the tool contract mid-benchmark and gives the model a feedback signal it
   did not have in the recorded runs, so post-change scores are not comparable.

**Overlaps with:** A-3 (route 3 is the only monitored route), A-5 (invalid ids),
O (warnings are not in any machine-readable artifact), C (no warning count in the run
stamp).

**Open questions.** How often does an honest model actually mis-stamp? The one observed
warning gives a rate of roughly 1 per 89 visible `create_*` calls in the opus run
(`grep -o 'create_event\|create_todo' outputs/opus.md | wc -l` → 89 across the tools and
narration lines), but the tool trace is lossy so the denominator is a lower bound, and
same-day mis-stamps are not in the numerator at all. Nothing short of a state dump (phase 1 / O) can answer this
from an existing run.

---

## A-5 A wrong `email_id` silently deletes the object from every node pool, so an honest model loses full credit with no signal
Status: open
Severity: distorts-measurement
Cost to verify: needs-one-live-run

**What's wrong.** `_node_state` resolves the stamp to a node and drops any object whose
id is not in the corpus, and routes any object whose id belongs to another node into
*that* node's pool. A model that does exactly the right work and mis-stamps therefore
gets "(nothing matching created)" for the email it served correctly, and, if the wrong id
is a valid one, can additionally break an unrelated email's uniqueness or cancel check in
the node it landed in. This is a measurement error in the opposite direction from gaming,
and the log renders it identically to genuine under-action.

**Evidence.**
- `runner.py:137-146`:
  ```
  e = corpus.emails.get(o.email_id)
  return o if (e and e.node == node) else None
  ```
  Unknown id → `e is None` → the object is in no `NodeState` for any node, and
  `runner.py:507` `by_eid.get(i) == eid` also excludes it from every `TurnDelta`. The
  object exists in the store, is visible to the model via `list_events`
  (`mcp_app.py:64-67`), and is invisible to the grader in both channels.
- Same-node wrong stamp is harmless for op grading (same pool) but still breaks the
  sibling's no-action verdict; cross-node wrong stamp is the damaging one.
  33 of 167 emails have a same-day same-node sibling, 150 have a same-day cross-node
  sibling:
  ```
  PYTHONPATH=$PWD .venv/bin/python -c "
  from datetime import date
  from sb.schema import load_corpus
  from sb.scheduler import Levers, build_plan
  c=load_corpus('corpus'); p=build_plan(c,start_date=date(2026,6,1),seed=42,n_days=60,levers=Levers(1,5,7))
  d=[b for b in p.per_day if b]
  print('same-node sibling:', sum(1 for b in d for e in b if any(c.emails[o].node==c.emails[e].node for o in b if o!=e)))
  print('cross-node sibling:', sum(1 for b in d for e in b if any(c.emails[o].node!=c.emails[e].node for o in b if o!=e)))"
  # same-node sibling: 33
  # cross-node sibling: 150
  ```
- The failure this produces is indistinguishable in the log from doing nothing:
  `grader.py:168` sets `actual = "(nothing matching created)"` whenever `title_set` is
  empty, and `grader.py:173` reasons `no {obj_word} titled like "{kw}" was created`.
  Counts: 52 opus, 47 sonnet, 52 haiku (`grep -c 'nothing matching created'`).
- **Negative result, stated so it is not overstated.** I could find no observed instance
  of cross-node misattribution in any run. Every object listed in every duplicate-failure
  `actual` field across all three logs is topically native to the node it appears under
  (e.g. `outputs/opus.md:954-957` lists only Project Atlas objects under
  `project_atlas.dinner-cancel`; `outputs/opus.md:71` lists only Innovation-comp objects).
  The single observed misattribution (A-6) was *within* a node. The case for this finding
  rests on the monitor's blind spot (A-4) plus the exposure count, not on observation.
  Note the log cannot refute it either: an object misattributed into a node where it
  matches no keyword never appears in any `actual` field.
- Where the blind spot and the failure overlap: 51 of opus's 52 and 46 of sonnet's 47
  "(nothing matching created)" details fall on multi-email days, which are exactly the
  days where a wrong stamp is silent.
  ```
  for f in outputs/opus.md outputs/sonnet.md past/claude-haiku-4-5.md; do printf "%-30s " "$f";
    awk '/── day /{split($0,a," · "); n=a[3]+0} /nothing matching created/{if(n==1)s++; else m++} END{printf "single-email days=%d multi-email days=%d\n", s+0, m+0}' "$f"; done
  # outputs/opus.md               single-email days=1 multi-email days=51
  # outputs/sonnet.md             single-email days=1 multi-email days=46
  # past/claude-haiku-4-5.md      single-email days=0 multi-email days=52
  ```

**Why it matters for benchmark validity.** §4.3 of the brief lists four candidate causes
for the 52 "nothing matching created" cases and says the split is unmeasured. One of the
four (wrong-node stamp) is a category-A cause and is currently unmeasurable from the
artifacts by construction, so the split cannot be closed by log analysis alone. Worse,
this error mode penalises exactly the behaviour the benchmark wants: a model that does
the work and files it under the wrong header scores zero and looks lazy.

**Options.**
1. Grade against all objects regardless of stamp, using node membership inferred some
   other way (e.g. title/description match against the answer key only). Removes the
   honest-model penalty entirely; it also removes the only mechanism that keeps one
   node's objects out of another node's pool, so cross-node keyword collisions (G) get
   strictly worse.
2. Fall back to a day-scoped pool when a stamp does not resolve: an unresolvable object
   is graded against every email served that day. Recovers credit for typos without
   loosening the general rule; it creates a perverse incentive where an unstampable
   object is graded more generously than a correctly stamped one.
3. Reject unresolvable stamps at the store so the model sees an error and can retry
   (`store_app.py:126` currently accepts any string, while `:225-230` 404s on read).
   Turns a silent scoring loss into a visible tool error; it changes the tool contract,
   and a model that cannot recover from the error now loses the object entirely rather
   than losing only the credit.
4. Leave the rule and instrument it: dump final state with stamps (phase 1 / O) and
   report how many objects failed to resolve. Decides nothing, costs nothing beyond the
   O work, and is the only option that can produce the number this finding is missing.

**Overlaps with:** A-3 and A-4, G (the "nothing matching created" bucket is shared with
title and kind mismatch), O (a state dump is what would measure this),
§4.3 of the brief.

**Open questions.** The honest mis-stamp rate is the number this whole finding turns on
and it is currently unknown. Also open: `update_event` / `update_todo`
(`mcp_app.py:49-56`, `:80-86`) carry no `email_id` and cannot change one, so an object's
stamp is fixed at creation. Is that intended? It means a model that realises it filed
something under the wrong email has no tool to correct it.

---

## A-6 Worked case: one stale stamp in `outputs/opus.md` produced one false PASS and cost one earned PASS
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** The single attribution warning in the entire recorded corpus of runs is
not benign: it marks a legitimate model action that escaped the no-action check on one
email and then failed a second, unrelated email's cancel check. The two errors run in
opposite directions and net to zero on opus's headline score, which is why the anomaly is
invisible in the aggregate while both individual verdicts are wrong. This contradicts
§2.6's reading of the same warning as "a theoretical hole, not an observed exploit": it is
observed and score-changing, though it is not an exploit (the model was acting honestly).

**Evidence.** Sequence, all from `outputs/opus.md` and the reconstructed plan:
- Day 1 (`outputs/opus.md:16-19`): `Innovation-comp.sponsor-mixer-before-the-final` is
  served, has no ops, and passes:
  `✓ PASS  [2] Innovation-comp.sponsor-mixer-before-the-final  · served Mon Jun 01`
  / `actual    (nothing)` / `why       correctly took no action`.
- Day 4 (`outputs/opus.md:67`) the model creates an event stamped with that day-1 email:
  `warning  stale_email_id (object_kind=event, email_id=Innovation-comp.sponsor-mixer-before-the-final, title=Sponsor mixer with retail partners (optional appearance))`.
  Day 4's batch is a single email (`plan.per_day` index 3 is
  `['Innovation-comp.pitch-comp-is-on-locking-in-the-final-da']`), so this is not a
  sibling stamp; it is route 2 of A-3, action on a later day.
- The day-1 verdict is never revisited: `runner.py:503-512` writes `results[eid]` once,
  inside the day iteration. So an event created explicitly for a no-action email keeps a
  "correctly took no action" pass.
- Day 35 (`outputs/opus.md:689-692`):
  ```
  ✗ FAIL  [104] Innovation-comp.sponsor-call-no-longer-needed  · served Sun Jul 05
     ✗ expected  event ~"sponsor" cancelled
       actual    "Sponsor mixer with retail partners (optional appearance)" Tue Jul 14 6 PM (120m)
       why       should be cancelled, but 1 still on the calendar
  ```
  The only surviving object is the mis-stamped mixer. The model had correctly deleted the
  real sponsor call (`outputs/opus.md:664` still lists
  `"Intro call — prospective comp sponsor (with BizDev)"` on day 33; day 35's tools line
  at `outputs/opus.md:687` contains `delete_event`). By `grader.py:156-157`,
  `passed = len(title_set) == 0`, so removing the mixer flips this op to pass.
- Counterfactual arithmetic, three variants, so the causal claim is precise:
  - *Object had never existed*: opus [104] passes. **90 → 91.**
  - *No-action re-evaluated at end of run over all objects carrying the email's id*:
    opus [2] flips to fail. **90 → 89.** Combined with the previous variant, 90 → 90 with
    two different emails now correctly scored.
  - *Stamp corrected to the day-4 email*: no change. The mixer is in node
    `Innovation-comp` either way, so `_node_state` routes it identically. The stamp is
    what made the error *visible*, not what caused it.
- Blast radius check, so the finding is not overstated: the same object also appears in
  `outputs/opus.md:664` under `[100] Innovation-comp.sponsor-wants-a-follow-up-call`,
  where it makes `found 2 matching`. Removing it leaves one object on `Tue Aug 25` against
  an expected `Wed Aug 12`, so that email still fails, with reason "on the wrong day"
  instead of "duplicate". Verdict unchanged, reason changed.
- Not reproduced by the other models: sonnet and haiku fail the same `[104]` email for a
  different reason, an earlier same-node object that also matches `"sponsor"`
  (`outputs/sonnet.md:689-692`, `past/claude-haiku-4-5.md:604-607`). So the +1 is specific
  to opus.

**Why it matters for benchmark validity.** This is the only ground-truth data point the
project has about attribution, and it shows the failure is two-sided and self-cancelling
in the aggregate. Any future check of the form "did the score move?" will report no
movement while two of 167 verdicts are wrong. It also shows that the damage travels: an
over-action that escapes its own email's check does not stay contained, because
`grader.py:152` pools by node cumulatively.

**Options.**
1. Treat this as the trigger `docs/POTENTIAL_GAMING.md:65-67` describes ("if either
   warning appears regularly") and harden now. One instance is not "regularly", so this
   reads the trigger loosely; against that, the instance is score-changing and the monitor
   cannot see the common case (A-4).
2. Treat it as a corpus problem instead: the mixer email is an FYI whose correct handling
   is genuinely ambiguous ("optional appearance"), so the answer key may be wrong rather
   than the grader. Cheapest fix and it addresses one email; it generalises to nothing and
   K would need to re-audit the other 55 no-action keys for the same ambiguity.
3. Use it as the seed case for phase 1.5's hand-graded sample: hand-grade the
   `Innovation-comp` sponsor thread end to end and see how many of its verdicts survive.
   Produces the honesty baseline the register already wants; it is manual work and one
   thread of 15.
4. Do nothing until the O work lands, on the grounds that a single instance cannot
   justify a rule change and a state dump would settle it properly. Avoids acting on n=1;
   defers the only category-A evidence that exists.

**Overlaps with:** A-3 (route 2), A-4 (the warning that surfaced it), G (the cumulative
node pool is what propagated it), K (`Innovation-comp` no-action keys).

**Open questions.** Was creating an "optional appearance" event for that email actually
wrong? The answer key says no ops, so the grader calls it over-action, but a human
assistant blocking an optional sponsor mixer is defensible. If the key is arguable, then
the false PASS in variant 2 above is not clearly false, and this case is evidence for K as
much as for A.

---

## Reporting back on the brief

### §2 claims I could not reproduce

- **§2.6's line anchor is wrong.** It cites `runner.py:472-475` for "grades a day's
  objects by splitting on the `email_id` the model stamped". Those lines are the retry
  and backoff block (`detail = (proc.stderr.strip() or ...)`, `except Exception as exc:`,
  `pause = _limit_reset_wait(...)`). The attribution split is at `runner.py:498-510`.
- **§3's anchor is wrong for the same reason.** "day loop, attribution split |
  `sb/live/runner.py:394-477`" stops 21 lines before the split it names.
- **§2.6's inference is unsound, not just its anchor.** "Across all three runs those
  warnings fired once (opus), so this is a theoretical hole, not an observed exploit"
  draws a conclusion the monitor cannot support: `store_app.py:86-92` has no branch for a
  same-day sibling id, so the sibling-stamp case it is arguing about produces zero
  warnings by construction. And the one warning that did fire is an observed,
  score-changing instance of a *different* escape route (action on a later day), not a
  null result. See A-4 and A-6.

Everything else in §2 that I touched did reproduce: the three scores (90/91/98, and
`past/claude-sonnet-4-5.md` at 102/176 on the retired corpus), the §2.3 tallies for
"over-acted" (5/5/0) and "correctly took no action" (51/51/56), the 57-day / 167-email
plan at default levers, 15 nodes / 167 emails / 56 no-action emails, and
`past/claude-opus-4-8.md` at 0 bytes.

### §4 claims I believe are now partly established

- **§4.3 (the split of the 52 "nothing matching created" cases).** Two of the four
  candidate causes can now be constrained.
  - *Wrong-node stamp*: zero observed instances in any of the three runs. Every object in
    every duplicate-failure `actual` list is topically native to the node it is listed
    under, and the single confirmed misattribution (A-6) was within a node. This does not
    rule the cause out (an object misattributed into a node where it matches no keyword is
    invisible in the log) but it means there is currently **no positive evidence** for it,
    so a proportion attributed to attribution would be unsourced.
  - *Whole-day inaction*: ruled out for 49 of opus's 52 and 42 of sonnet's 47 cases, which
    fall on days where the tools line shows at least one `create_event`/`create_todo`.
    Because the trace is lossy (§2.7) this is a lower bound, i.e. the true "the model was
    not idle" count is at least this high. It does not prove the model acted on *that*
    email, only that it was not globally inactive that day.
  The residual mass therefore sits with G (title mismatch, kind mismatch), not with A.
- **§4.5 (phase 0 did not fix the convergence).** I can now name a mechanism for part of
  it: haiku's 8-point lead over opus is 5 points of no-action (56/56 vs 51/56) and 3
  points of acting emails, and haiku is the only run with a different `daily_max`. See
  A-1.
- **§4.1 (verdict agreement cannot discriminate).** Supported and sharpened: 51 of the
  ~100 grader-determined outcomes the brief refers to are specifically no-action passes,
  and I can now give the exact per-model counts.
