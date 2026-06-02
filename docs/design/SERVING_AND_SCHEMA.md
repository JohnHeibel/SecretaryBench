# SecretaryBench — Serving Model & Corpus Schema (v0, for discussion)

> Status: **DRAFT.** Third design doc, after `BENCHMARK_REDESIGN.md` (the why) and
> `ANSWER_KEY_GRAMMAR.md` (the when/what). This one is the **how it's served**: the
> dependency model, the scheduler, and the on-disk format authors write into. Authored
> 2026-05-30. **[OPEN]** = needs a decision.

---

## 0. The core simplification: one email-level DAG

Decided: intra-node and cross-node dependencies use the **same typed-edge model**. So we
don't reason about two graph levels — we **flatten to a single DAG whose vertices are
emails.** A "node" is just an organizational grouping: a shared cast of characters and a
shared set of emitted anchors. It carries no scheduling semantics of its own; all readiness
is computed from email-level edges.

```
node = { id, cast, emails[], (emitted anchors live on the emails) }
edge = email_A --type--> email_B    // "B depends on A"
```

Node-level edges are allowed as **authoring sugar** and compile to email-level edges (all
of A's emails become prerequisites of all of B's). The scheduler never sees nodes.

---

## 1. Edge types (this is the whole constraint model)

| Type | Meaning | Imposes ordering? | Imposes a deadline? |
|------|---------|-------------------|---------------------|
| `static` | B may rely on a non-date fact A established (a rule, a name, a count). The fact never expires. | **Yes** (A before B) | **No** — B may be served *arbitrarily* later. |
| `date` | B's answer references a **date anchor** A emitted (`@signing + 2w`). | **Yes** (A before B) | **Yes** — see §2. |

Two consequences worth stating plainly:
- **`static` edges are the long-span retrieval lever.** Spread A and B as far apart as you
  like (day 3 → day 90); the further apart, the more the model must *search* its inbox for
  A's fact. No feasibility cost.
- **`date` edges are self-limiting.** The anchor arithmetic creates a hard serve-by
  deadline, so they can't span arbitrarily far without becoming incoherent.

---

## 2. Derived serve-by windows (the "served when it should be" guarantee)

A `date` edge means B's answer resolves to a concrete date *relative to A's serve date*.
Example: A emits `@signing = serve(A) + 5d`; B's answer is `@signing + 2w`. The correct
event date is `serve(A) + 19d`. If B is served *after* that date, the model is being asked
to schedule a meeting in the past — incoherent.

So: **B's serve-by deadline = the earliest answer-date in B, resolved once A is served.**

- The deadline is unknown until A is served (it depends on `serve(A)`), but B isn't even
  *ready* until A is served — so the moment B becomes ready, its deadline is computable by
  the resolver. No chicken-and-egg.
- Window for B = `(serve(A), earliest_answer_date(B)]`. Width = the anchor arithmetic
  (here ~19 days). Authors widen it by using larger offsets; the feasibility check (§5)
  refuses corpora whose windows can't be met.

`static` edges attach **no** deadline.

---

## 3. Readiness predicate

An unserved email `e` is **ready on day `d`** iff:

1. every email in `e.depends_on` was served on a **strictly earlier day** (`< d`); and
2. if `e` has any `date` dependency: its serve-by deadline `≥ d` (window still open).

Strict prior-day (rule 1) means a dependency and its dependent never land in the same daily
batch — which removes all intra-day ordering complexity. **Independent emails — including
independent emails in the same node — may share a day.** A dependency *chain* of length L
therefore takes ≥ L days to fully deliver, which is realistic (emails arrive over time).

---

## 4. The scheduler: seeded EDF + weighted random filler

Deterministic given `(corpus, seed, levers)`; only the *model's* responses are
non-deterministic. **Not** a pre-baked fixed schedule — it's pseudo-random, re-derived each
run, reproducible on the same seed, and steerable by the difficulty levers.

```python
rng = RNG(seed)
anchors = {}          # name -> concrete date, filled as emitters are served
deadlines = {}        # email -> serve-by day, filled when an email becomes ready
for d in range(N_DAYS):
    ready = [e for e in unserved if is_ready(e, d, served, deadlines)]

    # 1. EDF: anything whose window is closing must go first (never miss a deadline)
    urgent = sorted([e for e in ready if deadlines.get(e, INF) <= d + URGENCY_HORIZON],
                    key=lambda e: deadlines[e])

    # 2. Filler: sample the rest to hit today's target load — the difficulty knob
    target = sample_daily_count(rng, levers)        # 1..5  [distribution OPEN — see §6]
    fill   = weighted_sample(rng, [e for e in ready if e not in urgent],
                             k=max(0, target - len(urgent)), levers=levers)

    for e in cap(urgent + fill, DAILY_MAX):
        render_body_tokens(e, serve_date=d)         # tokens -> concrete dates in the email
        record_emitted_anchors(e, d, anchors)       # {!signing=+5d} -> anchors["signing"]
        run_model_turn(e, serve_date=d)             # harness
        mark_served(e, d)
        compute_deadlines_for_newly_ready(e, anchors, deadlines)
```

- **EDF guarantees** no `date`-windowed email is ever stranded, as long as the corpus passed
  the feasibility check.
- **`URGENCY_HORIZON`** is how many days ahead of a deadline we start force-serving — a small
  buffer so windows never slip even when daily slots are contended.
- A priority queue keyed on deadline is the natural implementation of `urgent`.

---

## 5. Feasibility pre-check (fail loud, before any model runs)

At load time, statically validate the corpus:

1. **It's a DAG** — no dependency cycles.
2. **Anchors resolve** — every `@name` referenced has an ancestor that emits it (via a
   `date` edge), and every answer date traces to a token (the grammar linter).
3. **Deadlines are satisfiable** — simulate the schedule with *EDF only* (drop the random
   filler, respect deps + deadlines + `DAILY_MAX`). If any `date` email's window can't be
   met, the corpus is over-constrained → **reject at load with the offending email named.**

This turns "we forgot to serve an email in time" from a silent mid-run bug into an
author-time error message.

---

## 6. Difficulty levers (all seeded; this is the difficulty dial)

| Lever | Effect |
|-------|--------|
| `daily_count` distribution | How many emails/day (1–5). Raises concurrent load/ambiguity. **[OPEN: pure-uniform vs weighted — John deferring]** |
| `span_bias` | How aggressively the filler delays the descendant of a `static` edge → retrieval difficulty (search day-3 facts on day-50). |
| `distractor_ratio` | Share of no-action / FYI mail mixed in → bigger haystack, "don't over-act" pressure. |
| `urgency_horizon` | Buffer before `date` deadlines; mostly a safety knob, secondarily affects clustering. |
| `seed` | Reproducibility. Same seed + levers ⇒ identical serving, 100%. |

Difficulty axes from `BENCHMARK_REDESIGN.md §6b` map onto these: temporal complexity
(grammar, authored), dependency depth (graph, authored), dependency **span** (`span_bias`),
distractor density (`distractor_ratio`).

---

## 7. On-disk corpus schema

Plain JSON in version control. One file per node; emails carry their own edges and answers.

```jsonc
// corpus/nodes/henderson.json
{
  "id": "henderson",
  "cast": { "V": "Dana Ruiz (Vendor Legal)", "CEO": "you" },
  "emails": [
    {
      "id": "henderson.intro",
      "from": "V", "to": "CEO",
      "subject": "Henderson acquisition — kickoff soon",
      "body": "Hi — we're moving forward on Henderson. More to come.",
      "depends_on": [],
      "answer": { "ops": [] }                           // no action expected
    },
    {
      "id": "henderson.signing",
      "from": "V", "to": "CEO",
      "subject": "Signing date",
      "body": "The signing is locked for {!signing = +5d}.",   // emits @signing, renders a date
      "depends_on": [ { "email": "henderson.intro", "type": "static" } ],
      "answer": { "ops": [
        { "create": "signing", "kind": "event", "on": { "eq": "@signing" } } ] }
    },
    {
      "id": "henderson.kickoff",
      "from": "CEO_boss", "to": "CEO",
      "subject": "Kickoff",
      "body": "Once Henderson signs, get the project kickoff on the calendar for two weeks after.",
      "depends_on": [ { "email": "henderson.signing", "type": "date" } ],  // ← carries a deadline
      "answer": { "ops": [
        { "create": "kickoff", "kind": "event", "on": { "eq": "@signing+2w" } } ] }
    }
  ]
}
```

Cross-node edges look identical — a `depends_on.email` just names an email in another node's
file. Optional node-level sugar:

```jsonc
// in node B: shorthand that compiles to email-level edges
"node_depends_on": [ { "node": "henderson", "type": "static" } ]
```

Field notes:
- `id` is globally unique; the `node.local` convention keeps it readable and namespaced.
- Anchor emission lives in the **body token** (`{!signing=+5d}`) so it renders *and* records
  from one source. (An `answer.emits` form remains available for anchors not shown in text.)
- `answer` follows `ANSWER_KEY_GRAMMAR.md` exactly — this doc doesn't redefine it.
- `from`/`to` reference `cast` keys so identities stay consistent across a node.

---

## 8. What this implies for the rest of the build

- **MCP server** gains `search_inbox(query)` + `get_email(id)` over already-served emails
  (the retrieval channel). Already flagged in the redesign doc.
- **Resolver** (likely a repurpose of `engine.py`'s token logic) does three jobs: render
  body tokens, record emitted anchors, and compute answer dates + deadlines.
- **Grader** is **state-based** (snapshot store after each email; reconcile each
  obligation's `create`/`move`/`cancel` op) per the grammar spec — not the legacy diff grader.
- **Scheduler** is the new module replacing `flow_controller.py`.

Build order (unchanged recommendation): schema + resolver + scheduler + state-grader, proven
end-to-end on a hand-authored 3–5 node corpus, *then* scale authoring.

---

## 9. Open items

- **[OPEN]** `daily_count` distribution — uniform 1–5 vs weighted. (John deferring; lever
  exists regardless.)
- **[OPEN]** `DAILY_MAX` / `URGENCY_HORIZON` default values (tune empirically on the pilot).
- **[OPEN]** Node-level sugar expansion rule: "all-of-A before all-of-B," or "A's terminal
  emails before B's initial emails"? (Former is safer; confirm.)
- **[OPEN]** Should a node be allowed to depend on *itself across runs* / recurring nodes?
  (Probably out of scope for v1.)
- **[CARRIED]** No-action strictness (B2), binary vs partial metric (B3), reply/delegate
  schemas, timezone lock — from the grammar doc, still open.
