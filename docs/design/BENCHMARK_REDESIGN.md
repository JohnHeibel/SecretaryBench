# SecretaryBench — Redesign Brief (v0, for discussion)

> Status: **DRAFT FOR BACK-AND-FORTH.** This document operationalizes the current
> state of the project and frames the decisions we need to make before a major
> reset. Nothing here is committed. Sections marked **[PROPOSAL]** are my opening
> position to be argued with; sections marked **[OPEN]** are questions for John.
>
> Authored 2026-05-30 after a full read of the codebase, the dataset, and the
> Sprint 1–5 docs.

---

## 1. What this benchmark is supposed to be

An **academic temporal-reasoning benchmark** for LLM "secretary" agents. We hand a
model an inbox of emails, one at a time. From each email the model must take the
correct secretarial action — create a calendar event, create a to-do, reschedule
something, or correctly do nothing — at the **correct date and time**. We then
grade the model's calendar/to-do state against a hand-authored answer key.

The intended research contribution is **temporal reasoning under realistic,
relative time pressure**: "next Thursday," "the third Friday of November," "two
weeks after the contract is signed," "before the deadline in the prior email."
The benchmark should be able to dial difficulty up and down by controlling **what
information arrives, and in what order**, without rewriting the emails.

---

## 2. Current architecture (as built)

```
Emails.xlsx  ──loader.py──▶  Scenario/Email objects
                                   │
                          flow_controller.py  (shuffle 109 scenarios across 100 days,
                                   │            enforce per-chain email order, random 0/1/2-day gaps)
                                   ▼
                              engine.py  (per simulated day: resolve {date} tokens →
                                   │       serve email → run model turn → diff store state →
                                   │       grade diff → mark served → grade scenario at completion)
                                   │
              ┌────────────────────┼─────────────────────┐
              ▼                    ▼                     ▼
        harness/ (claude -p   mcp_server/ (15 FastMCP    app/ (FastAPI in-memory
        adapter; codex stub)  tools → HTTP)              store: calendars/todos/
                                                          events/scenarios/emails)
                                   │
                              grader.py  (parse "TC-/CC-/RS-/No action" criteria,
                                          match against store state by date+content)
```

### What's worth KEEPING (clean, decoupled, reusable)
- **`mcp_server/`** — thin FastMCP wrapper, 15 tools (calendar/event/todo/email CRUD),
  stateless, routes to the API over HTTP. Easy to repoint at a new backend.
- **`app/`** — FastAPI in-memory store with clean Pydantic models for calendars,
  events, todos. The data model (event = title/start/end, todo = title/due_date) is fine.
- **`harness/`** — genuinely good abstraction. One `HarnessAdapter` interface,
  one source of truth for the system prompt + MCP config + calendar bootstrap +
  stream-json parsing + token/compaction telemetry. `ClaudeCodeAdapter` works;
  `CodexAdapter` is a ready-to-fill stub. Session continuity (`--resume`) is wired.

### What is BROKEN / must be rebuilt
- **`Emails.xlsx` (the dataset)** — mislabeled, AI-generated, inconsistent. 109
  scenarios / 281 emails. **40 distinct date-placeholder syntaxes**; **29 scenarios**
  where the body's date token doesn't match the criterion's date token; 13 scenarios
  with stray `^` markers; 9 scenarios that are ungradeable free text. Quality ranges
  from professional to literal spam copy. **Not salvageable as an academic artifact.**
- **The serving model is rigidly linear.** `flow_controller` shuffles once and deals
  scenarios across fixed days; emails within a scenario must go in order; gaps are
  random-forward-only. There is **no dependency graph** — scenarios are independent,
  and you cannot use ordering as a difficulty lever.
- **The date contract is backwards.** Tokens like `{date+9}` live **inside the email
  body** and get string-substituted before the model sees them. Authors who didn't
  understand the tokens wrote mismatched/garbled ones, and the grader checks a token
  the email never actually expressed (see C19 gala above). This is the root cause of
  "a substantial portion of the dataset is wrong."
- **Grading↔email linkage is weak.** Grading is done by diffing store state before/after
  a turn, plus a model-supplied `scenario_id`. There is no strict, per-email contract
  for "this created object answers this specific email."
- **The grader is a string-prefix parser** retrofitted to tolerate messy criteria,
  rather than a clean evaluator of a structured answer key.

---

## 3. Goals of the reset (what "good" looks like)

These are John's stated requirements, made concrete:

- **G1 — Order-independent serving.** Any email can be served on (almost) any
  simulated day. The *only* constraints are dependency constraints (below). Day 5 vs
  day 50 must not change the correct answer.
- **G2 — No hardcoded literal dates.** Every date in an email is *rendered from a
  governed token* relative to the serve date (or a named ancestor anchor), so it reads
  as a concrete date but never ties the answer to a fixed calendar. No serve-independent
  date strings. (Refined in §4 after round 1 — tokens stay in emails; one grammar.)
- **G3 — Fully automatic parsing.** Every email and every answer is machine-parsable
  end to end; no human in the loop to interpret a criterion.
- **G4 — Strict, email-linked grading.** Given an email, there is a deterministic,
  automatically-checkable expected action, and a mechanism that links a model-created
  calendar/to-do object back to the email it answers. No "ungraded free text."
- **G5 — Scenarios-as-a-DAG.** Scenarios are nodes; edges are dependencies. A node's
  emails may reference people/places/facts established in ancestor nodes. The server
  may serve any node at any time **as long as all ancestors have already been served**.
  This makes ordering a controllable difficulty lever while keeping content coherent.
- **G6 — Handwritten content.** Emails are human-authored for an academic artifact.
- **G7 — Difficulty is a knob.** We can produce easy/medium/hard runs from the same
  corpus by varying spacing, interleaving, and how much dependency chaining is required.

---

## 4. The central design idea **[DECIDED — round 1]**

> Earlier draft proposed "no tokens in the email body." **Rejected by John, correctly.**
> Real emails contain dates; banning them is unrealistic. The bug was never that tokens
> existed — it was 40 ungoverned syntaxes plus an answer key that drifted away from them.

The actual contract:

> **One governed token grammar is the SINGLE SOURCE OF TRUTH. A token renders into BOTH
> the email body AND the answer key, from the same serve anchor, so they cannot drift.**

- The **author writes a token once**, e.g. `{d+2}` or `{nth_friday:+1}`.
- At **serve time** the token renders into the email body as a **concrete, real-looking
  date** — *"Wednesday, August 3rd"* — computed from the day the email is served. The
  email reads like a real email and shows real dates.
- The **same token** drives the **answer key**: the grader expects `serve_date + 2 days`,
  computed from the same anchor the rendered email used.
- A **linter** rejects (a) any token not in the documented grammar and (b) any answer-key
  entry that doesn't trace to a token in its email. Drift like C19 becomes impossible.

The only thing banned is a **hardcoded, serve-independent literal date** ("August 3rd"
typed as a fixed string). Every date in the corpus is *rendered from a token*, so it
looks absolute but is relative under the hood. That preserves G1/G2/G3 while keeping
emails realistic. (Truly-fixed real-world dates like holidays are a possible future
edge case but constrain serving order — avoid in v1.)

**Grading is against ground truth, per email** (decided, see §7-A3): each email's
expected action is computed from the canonical token values, independent of whatever the
model did on earlier emails. Errors do not compound.

### Named anchors — the cross-node mechanism **[DECIDED — round 1]**

How information in one scenario affects another, made checkable:

- An **ancestor node may emit a named DATE anchor** whose value is fixed the moment its
  emitting email is served. E.g. node A emits `henderson_signing = serve_date(A) + 5d`.
  (Refined round 2: **emissions are date-only.** A non-date fact like "90-minute meetings"
  is static — known at authoring time — so it's baked straight into the descendant's
  answer key, not emitted at runtime. Only a *date* depends on when the scheduler serves
  the ancestor, so only dates need the emission machinery.)
- A **descendant's answer key binds to that anchor by name**: `@henderson_signing + 14d`.
- The **DAG edge guarantees** the ancestor (and thus the anchor's value) was served before
  the descendant is graded, so the grader can always resolve it from a run-wide anchor
  table. The model's job is to have *remembered* the ancestor fact and done the arithmetic.

**Worked cross-node examples** (the templates authors will follow):

1. **Relative-to-ancestor-date (flagship temporal case).**
   A (Legal): "Henderson signing locked for `{d+5}`" → emits `henderson_signing`.
   B (later): "Kickoff two weeks after the signing." → answer `@henderson_signing + 14d`.
2. **Fact that reshapes the answer (not a date).**
   A (HR): "Client meetings are now 90-min blocks." → emits `client_meeting_duration=90m`.
   B: "Meet `{next_thursday}` at 2?" → answer: event 2:00–3:30 (model that missed A → 60m → fail).
3. **Cadence emitted once, instantiated later.**
   A (HR): "Weekly compliance training every Tuesday 10am starting next month." → emits rule.
   B: "Block the first month of trainings." → answer: 4 events, consecutive Tuesdays 10–11.
4. **Constraint / negative dependency.**
   A: "In Tokyo the week of `{d+9}`, no meetings." → emits blackout window.
   B: "Find time for the review next week." → answer must avoid the blackout (or flag).

### DAG semantics **[DECIDED — round 1, intra-node order still OPEN]**
- A **node** = a small bundle of emails sharing a cast and a fact-set (e.g. "Henderson
  acquisition"). 1..N emails. May declare named anchors it emits.
- A **directed edge A → B** = "B may reference facts/anchors first established in A ⇒ all
  of A's emails must be served before any of B's." Content dependency that imposes ordering.
- **Serving rule:** any topological order; spacing and interleaving within it are free and
  are the difficulty knob. Intra-node email order: still **[OPEN — §7-C2]**.

---

## 5. Proposed target architecture **[PROPOSAL]**

```
corpus/  (handwritten, version-controlled, plain-text/JSON — NOT Excel)
  nodes/<node_id>.{md|json}        ← emails w/ governed tokens + emitted named anchors
  graph.json                        ← nodes + dependency edges
  answers/<node_id>.json            ← structured answer key per email (tokens/anchors)
        │
   loader (validates DAG; validates answer schema; lints tokens against the grammar
        │   and checks every answer traces to a token in its email — no drift)
        │
   scheduler (topological-order-respecting; emits a serve plan with arbitrary
        │      spacing/interleaving per a difficulty profile)
        │
   engine (per served email: inject real serve-date → run model turn via harness →
        │   capture created/modified objects → attribute to email → grade)
        │
   grader (evaluate structured answer-key temporal expressions against serve anchor;
        │   deterministic, no string-prefix guessing)
   ┌────┴───────────────┐
 harness/ (KEEP)     mcp_server/ + app/ (KEEP, lightly adjusted)
```

What changes vs. today: **dataset format, loader, scheduler (was flow_controller),
grader, and the engine's date logic are rewritten.** Harness, MCP server, and store
survive largely intact.

---

## 6. Hardest open problems (flagging risk early)

1. **Answer-key expression language.** We need a small, closed grammar of temporal
   expressions rich enough for real secretarial tasks (relative offsets, nth-weekday,
   "before deadline X," time-of-day, durations, ranges) but closed enough to evaluate
   deterministically. This is the core IP of the benchmark and the thing most likely
   to balloon.
2. **Email→object attribution for grading.** How do we know *which* created event
   answers *which* email, when the model may batch, reorder, or partially act? Options
   in §7.
3. **DAG authoring burden.** Handwriting a coherent, dependency-linked corpus with a
   validated answer key for every email is real work. Scope and tooling matter.
4. **Cross-node temporal references** without literal dates — the answer key must be
   able to point at "the signing date established in node A," which means ancestor
   nodes have to *emit named facts/events* that descendants' answers can bind to.

---

## 6b. Difficulty model — reasoning vs. memory **[DECIDED — round 1]**

> John's first instinct was "more emails than fit in context." Walked back (small
> context model), and we go further: **context-overflow is the wrong primary difficulty
> axis** — it confounds two capabilities and muddies the research claim.

Two distinct capabilities are easy to bundle by accident:
- **(a) Temporal reasoning** — given the needed facts in-context, compute the right
  datetime (nth-weekday, anchor arithmetic, blackout avoidance, durations). ← the goal.
- **(b) Long-horizon memory** — retain a fact from day 3 to use on day 50.

Flooding the context measures (b), and a bigger model then "wins" partly because it
forgot less, not because it reasons about time better — a muddy result for a *temporal
reasoning* paper. **So: make the reasoning hard, keep needed facts retrievable, and treat
memory-span as a separate labeled knob, not a confound baked into everything.**

The DAG + scheduler give four clean, taggable difficulty axes:
1. **Temporal complexity** of the answer: offset → nth-weekday → anchor-arithmetic →
   blackout-avoidance.
2. **Dependency depth**: how many ancestor anchors must be combined.
3. **Dependency span → a RETRIEVAL axis** (refined round 2). John's real target isn't
   "hold 200k tokens in context" — it's "force the model to *search past emails* to find
   the facts it needs, especially day 3 → day 50, reasoning across several retrieved
   emails at once." So the needed fact lives in an earlier email that has scrolled out of
   context; the model must recognize the gap, search the inbox, find the right email among
   distractors, and reason over it. The further the span, the more retrieval is forced.
   Run the same corpus at span≈2 ("fact still in context") vs span≈40 ("must search");
   **the gap between those scores is a headline result** and a clean retrieval ablation.
4. **Distractor density**: no-action / FYI mail between actionable ones (don't
   over-schedule, correctly defer) — *also* raises retrieval difficulty (more haystack).

> **Tool-surface requirement (new):** this demands a **searchable inbox of already-served
> emails** — `search_inbox(query)` and `get_email(id)` tools — which the MCP server does
> not fully expose today. Add it. We deliberately do **not** stuff context; retrieval is
> the channel for old facts. (Context-overflow as the difficulty engine is rejected — it
> confounds reasoning with memory; a multi-novel inbox is also just unrealistic.)

Tag every node with this difficulty vector and **report scores per tier** → a curve
across models, not one brutal wall. Goal is difficulty that *discriminates* (spread,
floor, headroom), not maximal failure.

**Scale (round-1 estimate):** ~100 days × 1–5 emails/day ≈ a few hundred emails across
~50–150 nodes; most cheap (no-action / context / single-create), with the actionable
temporal puzzles concentrated. Authorable. Build format + linter + a full end-to-end run
on ~5 nodes **first**, then scale authoring.

---

## 7. Decisions so far + still-open questions

### DECIDED in round 1
- **A1 — token contract:** tokens stay in emails; one governed grammar; single source of
  truth renders into both email and answer key; linter enforces. (§4)
- **A3 — cross-references graded against ground truth**, per email, errors don't compound. (§4)
- **B1 — attribution:** model passes an opaque `email_id` on every create call.
- **C1 — node/edge definitions** as in §4 DAG semantics.
- **Difficulty model** per §6b; memory-span is a separate ablation, not a baked-in confound.
- **D2 — clean break on a new branch** (assumed unless John objects).
- **Corpus format — plain-text/JSON in version control, NOT Excel** (assumed).

### STILL OPEN — next round
- **A2 — grammar scope.** Need John's gnarliest *real* example. Candidate primitives to
  confirm/cut: relative offsets, nth-weekday-of-month, "by/before deadline", time-of-day,
  durations, calendar-vs-business days, "the week of …", date ranges. This bounds the
  grader and is the next thing to spec.
- **B2 — no-action strictness.** Any created object on a no-action email = fail? Or only
  certain object types? Keep spam/FYI as a class — confirm grading rule.
- **B3 — binary vs partial credit.** Strict binary per email, or partial (right action /
  wrong time)? Drives the headline metric (E1). My lean: binary primary metric + a
  diagnostic "right-action-wrong-time" secondary, for error analysis.
- ~~C2 — intra-node email order~~ **DECIDED:** same typed-edge model as cross-node ("DAG
  inside a DAG"); flattened to one email-level DAG. See `SERVING_AND_SCHEMA.md`.
- **C4 — temporal edges:** do any edges encode "B ≥ N days after A" as a *serving*
  constraint, or is all timing carried inside email text + answer key? (My lean: keep
  edges purely content/order; carry all timing in tokens/anchors — simpler scheduler.)
- **D1 — salvage check:** the `engine.py` token resolver is decent — repurpose it as the
  answer-key evaluator rather than rewrite from scratch? Confirm keep-list otherwise.
- **D3 — store model:** in-memory single-calendar fine, or per-scenario isolation /
  persistence for analysis?
- **E1 — headline number** for the paper. (Most likely: "% emails actioned at correct
  datetime," sliced by difficulty tier + by dependency-span.) Confirm.

---

## 8. My recommendation in one line

The reset is real but **bounded**: keep the three clean infrastructure layers
(harness, MCP, store — plus a new searchable-inbox tool), throw away the
dataset/loader/scheduler/grader, and let the **token & answer-key grammar** be the thing
we design first and most carefully — everything else (DAG, scheduler, attribution,
difficulty tiers) hangs off getting that contract right.

→ Design docs, in reading order:
1. **`BENCHMARK_REDESIGN.md`** (this) — the why, the goals, the keep/rebuild split.
2. **`ANSWER_KEY_GRAMMAR.md`** — the token & answer-key language (the when/what).
3. **`SERVING_AND_SCHEMA.md`** — the dependency model, scheduler, and on-disk format
   (the how-it's-served). Decided: one email-level DAG; seeded EDF + weighted filler;
   date-edges carry serve-by windows, static-edges are the long-span retrieval lever.

Remaining open items live at the bottom of each doc; they're the round-2 agenda. The
buildable next step is the schema + resolver + scheduler + state-grader on a 3–5 node pilot.
