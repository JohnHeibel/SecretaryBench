# SecretaryBench — Token & Answer-Key Grammar (v0.2, verb-based)

> Status: **DRAFT.** The closed, deterministic language for *when* each email's correct
> action falls due. Companion to `BENCHMARK_REDESIGN.md`. Revised 2026-05-30 after John's
> call to cut idiomatic bloat; **2026-06-01 the answer key moved from `expect`/`count` to
> `create`/`move`/`cancel` verbs on named obligations (ADR 0001).** **[OPEN]** = needs a decision.
>
> **Answer key in one line:** an email's `answer` is a list of `ops`, each a verb on a named
> obligation — `{ "create": "kickoff", "kind": "event", "on": { "eq": "@signing+2w" } }`. The
> name is the obligation's identity *and* a node-scoped anchor, so a later `move`/`cancel`
> references it by name. There is no `count` field; "exactly one" is inherent to an obligation.
>
> **DECIDED (day-level grammar):** times-of-day and durations are **out**. Tokens resolve to
> whole **days**; the grader checks the *day* an event/todo lands on, nothing finer. Times may
> still appear as plain prose in an email ("at 2pm") but are never tokens and never graded.
> Removed since the earlier draft: the `@HH:MM` time-attach, the `DATETIME` value type, the
> `duration` answer field, and the `exact_time` tolerance. Within-day overlap/conflict is
> therefore not represented — reintroduce times only if conflict-avoidance becomes a target.

---

## 0. The one principle that keeps this small

**Only dates vary with serve order. Everything else is static.**

A 90-minute duration is 90 minutes whether the email lands on day 5 or day 50. A name
("Sarah is the new EA"), a rule ("every Tuesday"), a count ("three to-dos") — all known at
authoring time, none shift with the serve anchor.

So:
- **Tokens encode dates only.** They render to a concrete date in the email and resolve to
  the grader's expected date — one source, no drift (the C19 bug becomes impossible).
- **Everything non-temporal is plain prose** in the email, plus a **static field** in the
  answer key (the obligation `name`, its `kind`, the `match` keywords). No token, no emission.
- **The only dynamic cross-node dependency is date-on-date** (kickoff = signing + 2w, where
  the signing date isn't known until the scheduler serves the ancestor). That, and only
  that, needs a **named date emission**.

Everything in the email is converted to natural language before the model sees it; tokens
are an authoring + grading device, never shown raw.

---

## 1. Value types (intentionally few)

| Type | Meaning | Resolved example |
|------|---------|------------------|
| `DATE` | a calendar day | 2026-08-03 |
| `INTERVAL` | a date span | the week of 2026-08-03 |

Non-temporal values (`DURATION`, counts, names, clock times) are **static prose or
answer-key fields, not token types.** Grading granularity follows the token: a
`DATE` token matches the day, and an `INTERVAL` token matches containment in a
date span.

---

## 2. Anchors

| Anchor | Meaning |
|--------|---------|
| `serve` | the date the email was delivered — its "now" (default anchor) |
| `@name` | a **date/interval** emitted by an ancestor email |

A named anchor is frozen when its emitting email is served and stored in a run-wide table.
The DAG guarantees the emitter precedes any referencing email, so `@name` always resolves.
**Anchors are date-only** — there is no "fact" anchor (a non-date fact is static, so it's
baked straight into the descendant's answer key).

---

## 3. Token grammar (dates only)

Same compact form serves as the in-email token `{ … }` and the answer-key expr.

```
expr      := base (offset)*
base      := "serve" | "@" NAME | selector
offset    := ("+"|"-") INT unit
unit      := "d" (calendar days) | "bd" (business days) | "w" | "m" | "y"
selector  := "next:" WD ("from" expr)?  // next WD strictly after serve, or after expr's date
           | "this:" WD ("from" expr)?  // WD in serve's (or expr's) Mon–Sun week
           | "nth:" N "," WD "," monthref   // Nth WD of a month (N = 1..5 or "last")
           | "dom:" D "," monthref          // the Dth day-of-month
           | "week_of:(" expr ")"           // INTERVAL: Mon–Sun containing expr
           | "month:" monthref              // INTERVAL: a whole month
monthref  := "0m" | ("+"|"-") INT "m"   // relative to serve's month
WD        := MON|TUE|WED|THU|FRI|SAT|SUN
```
(No time-of-day: tokens resolve to a DATE or INTERVAL only.)

| Token | Resolves to |
|-------|-------------|
| `{+9d}` | serve + 9 calendar days |
| `{+3bd}` | serve + 3 business days |
| `{next:THU}` | next Thursday after serve |
| `{next:MON from @signing}` | the first Monday after the signing |
| `{this:FRI from @migration}` | the Friday of the week containing the migration |
| `{nth:3,FRI,+1m}` | 3rd Friday of next month  ← the C19 gala, correct |
| `{nth:last,FRI,0m}` | last Friday of this month |
| `{dom:25,0m}` | the 25th of this month |
| `{week_of:(serve+1w)}` | next week's Mon–Sun interval |
| `{@signing+2w}` | two weeks after the ancestor's signing date |

**Date emission** (ancestor declares a named date anchor), inline or in node metadata:

```
{!signing = +5d}              // emits @signing = serve+5d, and renders the date in the email
{!blackout = week_of:(+9d)}   // emits an INTERVAL
```

---

## 4. Predicates — the expected date as a value, set, or bound

A date slot is matched by one of:

| Form | Meaning | Use case |
|------|---------|----------|
| `eq: <expr>` | exact (default) | "Tuesday at 3" |
| `in: <interval>` | falls within | "sometime next week" |
| `by: <expr>` | on or before (≥ serve) | "finish by Friday" (todo deadline) |
| `in: <interval>, not_in: @anchor` | within a window, avoiding a blackout | the Tokyo case |
| `any_of: [<expr>, …]` | matches any listed | "Mon, Tue, or Wed work" |

A predicate is the value of an op's `on` field. Cardinality never appears: one obligation is
one object, so `create`/`move` mean "exactly one match on this day" and `cancel` means "none."
`in`/`not_in`/`any_of` replace the old `or`-clause string hackery with clean set-membership.

---

## 5. Answer-key entry — verbs on named obligations

An email's `answer` is a list of `ops`. Each op is one verb (`create` / `move` / `cancel`)
on a named obligation. The verb's value *is* the obligation name.

```jsonc
{
  "ops": [
    {
      "create": "kickoff",              // verb: create | move | cancel; value = obligation name
      "kind": "event",                  // event | todo  (create only)
      "on": { "eq": "@signing+2w" },    // the date predicate (a DAY); create/move only
      "match": ["kickoff"],             // OPTIONAL title keywords; defaults to [name]
      "tolerance": "exact_day"          // OPTIONAL: exact_day (default) | within:Nd
    }
  ]
}
```

- **The name does three jobs.** It is the obligation's identity (so `move`/`cancel`
  reference it later), the default title `match`, and — when a sibling op references `@name`
  — a node-scoped anchor equal to the obligation's date. Authors write the bare name; the
  loader qualifies it internally (`__obl_<node>__<name>`) so two nodes can both have a
  `kickoff`.
- **`match` defaults to `[name]`.** Override it only when the natural calendar title differs
  from the slug (e.g. `"create": "filing", "match": ["HSR"]`). All grading is fuzzy: the
  model's object matches if its title contains every keyword.
- **No-action / FYI / bait email** → `ops: []`. The turn must create nothing attributable
  to that email in the current day-loop grader.
- **Todo with a deadline** → `"create": "...", "kind": "todo", "on": { "by": "@x+5bd" }`.
- **Reply / delegate** verbs still TBD. **[OPEN]**

### 5.1 Move / cancel — reference an obligation by name

These grade on **final state**, not the model's mechanics — a model may reschedule by
`update_event` *or* by `delete`+`create`, and both score identically. Express the *intended
end state*, and let the obligation's identity catch leftovers:

```jsonc
// email A — "Get the board meeting on the calendar for next Thursday."
{ "ops": [ { "create": "board", "kind": "event", "on": { "eq": "next:THU" } } ] }

// email B — "Move the board meeting to the following Monday."
//   No depends_on needed: the edge to A derives from `move: board`. And @board is A's date,
//   so you don't repeat it. Exactly-one catches "created a new one, forgot the old" (a
//   double-booked calendar fails).
{ "ops": [ { "move": "board", "on": { "eq": "@board+4d" } } ] }

// email C — "Cancel the board meeting."   no `on`, no `kind` — just the name.
{ "ops": [ { "cancel": "board" } ] }
```

`move`/`cancel` inherit the obligation's `kind` and `match` from its `create`, and auto-derive
their dependency edge to the creating email (a `date` edge when the op references an anchor,
else `static`). That is the ergonomic win over the old model: name the thing once, then act
on it by name.

---

## 6. Recurrence — **[NOT IMPLEMENTED]**

A recurring series ("training every Tuesday for four weeks") would be a single `create` op
with a static cadence block resolving to N day-grained events. The grammar sketch:

```jsonc
{ "create": "training", "kind": "event", "on": { "eq": "month:+1m" },
  "recurrence": { "every": "TUE", "count": 4 } }   // start token + static cadence
```

No times-of-day (those are prose, not graded — see the header). Not built yet; the schema
does not parse `recurrence`. Revisit if a recurring obligation becomes a target.

---

## 7. Cross-node templates (revised — no fact emissions)

1. **Date-on-date (a body emission the descendant references).**
   A: *"Signing locked for `{!signing=+5d}`."*  B: *"Kickoff two weeks after the signing."*
   → B op: `{ "create": "kickoff", "kind": "event", "on": { "eq": "@signing+2w" } }`, with a
   `date` edge A→B (the serve-by window).
2. **Obligation referenced later in the same node (no hand-wiring).**
   A: *"Get the kickoff on the books."*  B: *"Push the kickoff back three days."*
   → B op: `{ "move": "kickoff", "on": { "eq": "@kickoff+3bd" } }`; the edge to A and the base
   date both derive from the name.
3. **Avoidance.**  A: *"Out the `{!blackout=week_of:(+9d)}`."*  B: *"Review next week."*
   → B op: `{ "create": "review", "kind": "event",
   "on": { "in": "week_of:(serve+1w)", "not_in": "@blackout" } }`, with a `date` edge A→B.

---

## 8. Grading semantics

- **Grade on cumulative final state, not the per-email diff.** A reschedule done as
  `update_event` and one done as `delete`+`create` must score identically, so the grader
  checks the calendar/todo *state* after the email, not the edit the model made.
- **Reconcile each obligation against the node's state.** For a `create`/`move` op, find the
  objects whose title contains the obligation's `match` keywords: there must be exactly one,
  and it must fall on the `on` day. For `cancel`, there must be none. Uniqueness is inherent
  to an obligation, so a reschedule that left a stale duplicate fails — no `count` needed.
- Match a model object to an obligation by `match` keywords within the email's node. Resolve
  date tokens against `(serve_date, anchor_table)` — **ground truth**, independent of the
  model's earlier actions.
- Tokens resolve to a DATE (or INTERVAL) → **day equality**; `within:Nd` relaxes the window.
- **No-action / FYI / bait email** (`ops: []`) → the model must create nothing attributable
  to that email. The live runner now works by day, but created objects still carry an
  `email_id`, so bait emails stay discriminating.
- Per-email result **binary** + a "right-action-wrong-day" reason in the details. **[OPEN B3]**

---

## 9. Authoring style guide

The grammar is exact; English isn't. Canonical phrasing, one per token, prevents drift:
`{next:THU}` ⇄ "next Thursday" (always the coming one); `{nth:3,FRI,+1m}` ⇄ "the third
Friday of next month"; `{this:FRI}` ⇄ "this Friday." A linter validates tokens and checks
every date answer traces to a token; it can't catch ambiguous prose, so the style table is
human discipline shipped with the authoring kit.

---

## 10. Open items

- **[OPEN]** Anything real emails need that a date token can't express? (multi-day events;
  fuzzy "morning/afternoon"; quarter/fiscal dates; date *ranges as the deliverable*.)
- **[DECIDED B2]** No-action strictness: empty `ops` ⇒ the turn must create nothing
  attributable to that email in the current day-loop grader (§8).
- **[OPEN B3]** Binary vs partial-credit headline metric.
- **[OPEN]** Reply / delegate verbs (`create`/`move`/`cancel` now implemented in §5.1).
- **[DONE]** Weekday selectors relative to a *named anchor* ("the Monday after the
  signing"). Implemented as an optional `from <expr>` clause on `next:`/`this:` (e.g.
  `next:MON from @signing`, `this:FRI from @migration+1w`); with no `from`, behavior is
  unchanged (serve-relative).
- **[OPEN]** Timezone & locale lock; default valid-slot rules for `in:` answers.
- **[IMPLEMENTATION NOTE]** Grader is **state-based** (snapshot the calendar/todos after each
  email and reconcile each obligation's `ops` against it), not diff-based. `sb/grader.py`.
