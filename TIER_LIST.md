# Email difficulty tiers — the authoring playbook (T1–T3)

This is the spec for what kind of email to write at each difficulty, so the work
can be divided up and so the test set has a real easy → hard gradient. Read
`RECAP.md` for the system and `ANSWER_KEY_GRAMMAR.md` for the grammar; this doc is
only about *how hard to make each email and why.*

## The one thing to internalize first

We are NOT testing date arithmetic. We learned (3 live runs) that the floor model
already nails business-day chains, reschedules, and cross-node anchors — synthetic
date-math difficulty does **not** separate models. What separates them is:

1. **Retrieval span** — how long ago the needed fact was set. Recent (in context) is
   easy; many emails/days ago (scrolled out, must `search_inbox`) is hard.
2. **Task recognition** — is it even obvious an action is required? The real failure
   mode is **under-action**: the model reads an email, thinks "FYI," and misses that
   it implies an obligation or a constraint. This is where the discrimination lives.

So difficulty climbs along *those* axes, not "make the date math gnarlier."

```
            T1 ─────────────▶ T2 ─────────────▶ T3
  span      same email        a few days back     weeks back (out of context)
  recall    none              one hop, recent     multi-hop / search required
  task      obvious           obvious-ish         ambiguous / implied / constrained
  failure   ~never            occasional          UNDER-action, missed constraint
```

---

## T1 — Easy (the floor / sanity layer)

**Purpose:** confirm the model can do the basic loop at all. Almost everyone should
pass these. They anchor the low end and catch totally broken models.

- **Self-contained.** Everything needed is in *this* email. No earlier fact required.
- **Date from `serve` only.** `serve+Nd`, `next:THU`, `dom:25,0m`. No `@anchor`.
- **Obvious action.** The email plainly asks for a calendar event or a to-do — or is
  plainly a no-op FYI.
- **No dependencies.** Stands alone in its node.

```
  EMAIL  (henderson.intro style)
  ───────────────────────────────
  From: Dana Reed <dana@henderson.com>
  Subject: Quick sync next Thursday

  "Hey — can you put a 30-minute intro sync on the calendar for next Thursday?
   Nothing to prep, just want to say hello."

  ANSWER KEY
  ──────────
  { "ops": [ { "create": "intro sync", "kind": "event", "on": { "eq": "next:THU" } } ] }
```

Also T1: the **clean no-action** email (a newsletter, a "thanks, received!" reply) →
`{ "ops": [] }`. Keep these *plainly* non-actionable at T1; the *tricky* no-ops belong
in T3.

**Keep it discriminating?** You can't, much — that's fine. T1 is the baseline, not the
ranking layer. Just don't let T1 dominate the set or the whole benchmark floats high.

---

## T2 — Medium (one hop, recent)

**Purpose:** test recall + cross-day reasoning while the fact is still plausibly in
context. This is the bulk of an interesting corpus.

- **One dependency, recent.** References a date or fact set in an *earlier* email in the
  same node, a few days back — close enough that memory *might* still hold it.
- **One anchor + offset**, a business-day count, or a **reschedule** of an earlier
  obligation. (`move`/`cancel` live here naturally.)
- **Action is clear once you recall the fact** — the email tells you to do something;
  the only work is fetching the one thing it points at.

```
  EARLIER EMAIL (sets the anchor)
  ───────────────────────────────
  "Signing is locked for {!signing=+5d}."        → publishes @signing

  THIS EMAIL  (henderson.kickoff style)
  ───────────────────────────────
  "Once the contract's signed, get the project kickoff on the books for two weeks
   after the signing."

  ANSWER KEY
  ──────────
  { "ops": [ { "create": "kickoff", "kind": "event", "on": { "eq": "@signing+2w" } } ] }
  depends_on: a date edge to henderson.signing
```

Reschedule flavor (also T2):

```
  "Push the kickoff back three business days."
  { "ops": [ { "move": "kickoff", "on": { "eq": "@kickoff+3bd" } } ] }
```

**Keep it discriminating?** The lever is the **span**: bury the anchor a few more days
back, or add 1–2 unrelated emails between the setup and the payoff, so recall isn't
free. Don't make the date math harder — make the *fact farther away.*

---

## T3 — Hard (the discriminating layer — where the benchmark earns its keep)

**Purpose:** separate strong models from weak ones. Every T3 email should stack **at
least two** of the three hardeners below. This is where you spend your best authoring
effort.

### Hardener A — Long retrieval span (fact is out of context)
The needed fact was set *many* emails/days ago — far enough it has scrolled out of the
model's window, so it MUST `search_inbox` to recover it. A strong model searches and
recovers; a weak one hallucinates or gives up.

### Hardener B — Multi-fact combination (cross-node / cross-email constraints)
The correct action needs a fact from email X *and* a constraint from email Y — and Y
isn't mentioned in the email you're reading. Classic: a **policy/blackout** email set
weeks ago that silently governs this action.

```
  WEEKS AGO (a policy email, different topic)
  ───────────────────────────────
  "Reminder: the office is closed and no external meetings the {!blackout=week_of:{+30d}}."

  THIS EMAIL
  ───────────────────────────────
  "Let's get the client review on the calendar sometime next week."

  ANSWER KEY
  ──────────
  { "ops": [ { "create": "review", "kind": "event",
               "on": { "in": "week_of:(serve+1w)", "not_in": "@blackout" } } ] }
```

The trap: the review email says nothing about the blackout. A model that doesn't recall
and apply the policy will schedule *into* the blackout — and fail. That's a real,
human-meaningful miss.

### Hardener C — Ambiguous task recognition (the under-action trap)
The email does NOT say "schedule X." It *implies* an obligation, or it looks like an FYI
but isn't. The model has to *recognize* that action is required.

```
  EMAIL  (looks like an FYI, actually implies a deadline)
  ───────────────────────────────
  "Heads up — legal says the HSR filing window closes 30 days after signing, and
   we absolutely cannot miss it."

  ANSWER KEY
  ──────────
  { "ops": [ { "create": "HSR filing", "kind": "todo", "on": { "by": "@signing+30d" } } ] }
```

No "please add a task" anywhere. The model must infer the to-do from the consequence.
This is the single most discriminating pattern we have — author lots of these.

### The T3 anti-pattern (don't do this)
A no-action **bait** email whose bait is *too obvious* ("FYI, no action needed" stamped
on it). That doesn't trap anyone. Real ambiguity means a reasonable person could go
either way until they think it through. If it's labeled, it's T1.

```
  T3 GENERATOR = pick ≥2:  [ span: fact far back ]
                           [ combine: needs a second email's constraint ]
                           [ recognize: action is implied, not stated ]
```

---

## How to use this when authoring

- **Tag every email** with its intended tier (T1/T2/T3) so the set can be balanced and
  the report can show score-by-tier. (Ask the webapp to carry a `tier` field, or note it
  in the email metadata.)
- **Aim for a gradient, not a wall.** A healthy set might be ~30% T1, ~40% T2, ~30% T3.
  Too much T1 floats every model high; too much T3 with no gradient just looks like noise.
- **The score that matters is T3.** T1/T2 prove the harness works; T3 is what ranks
  models. If two models tie on T1+T2 and split on T3, the benchmark did its job.
- **Keep it human-written.** Generated needles make the report circular ("we can't claim
  we test reasoning if a template stamped the reasoning"). The ambiguity in Hardener C
  especially has to come from a human writing natural, slightly-under-specified email
  prose — that's the part a generator can't fake.

## Quick reference

| | T1 Easy | T2 Medium | T3 Hard |
|---|---|---|---|
| Fact location | this email | a few days back | weeks back / out of context |
| Recall | none | one hop, recent | search required / multi-hop |
| Date source | `serve` only | one `@anchor` | anchor + constraint from elsewhere |
| Task clarity | explicit | explicit | **implied / ambiguous** |
| Verbs | `create` / `[]` | `create`/`move`/`cancel` | any, often `by`/`in`+`not_in` |
| Discriminates? | no (baseline) | a little (via span) | **yes — this is the point** |
| Effort to author | low | medium | **high — spend it here** |
