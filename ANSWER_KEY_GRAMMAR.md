# SecretaryBench — Token & Answer-Key Grammar (v0.1, slimmed)

> Status: **DRAFT.** The closed, deterministic language for *when* each email's correct
> action falls due. Companion to `BENCHMARK_REDESIGN.md`. Revised 2026-05-30 after John's
> call to cut idiomatic bloat. **[OPEN]** = needs a decision.
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
  answer key (`duration: 90m`, `count: 3`, `title_match: ["kickoff"]`). No token, no emission.
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
| `DATETIME` | a day with a time | 2026-08-03T14:00 |
| `INTERVAL` | a date span | the week of 2026-08-03 |

Non-temporal values (`DURATION`, counts, names) are **static answer-key fields, not token
types.** Grading granularity follows the token: a `DATE` token → match the day; a
`DATETIME` token → match day **and** time.

---

## 2. Anchors

| Anchor | Meaning |
|--------|---------|
| `serve` | the date the email was delivered — its "now" (default anchor) |
| `@name` | a **date/datetime/interval** emitted by an ancestor email |

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
{!blackout = week_of:{+9d}}   // emits an INTERVAL
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

(Cardinality is **exactly one by default** — an obligation is a single thing. The static
`count` field is only for exceptions: `0` = must not exist, `N` = genuinely several.)
`in`/`not_in`/`any_of` replace the old `or`-clause string hackery with clean set-membership.

---

## 5. Answer-key entry (tokens for dates, static fields for the rest)

```jsonc
{
  "email_id": "henderson-kickoff",
  "emits": { "signing": "+5d" },            // date anchors this email establishes (if any)
  "expect": [
    {
      "action": "create_event",             // create_event | create_todo | reschedule | reply | delegate
      "title_match": ["kickoff"],           // static: keyword(s), case-insensitive
      "start": { "eq": "@signing+2w" },      // the only token-driven field (a DAY)
      "tolerance": "exact_day"               // exact_day (default) | within:Nd
      // cardinality defaults to exactly one — omit it. Set `count` only for 0 (cancel) or N.
    }
  ],
  "forbid": []                               // no-action emails: forbid all creates  [OPEN B2]
}
```

- **No-action email** → `expect: []`, `forbid` all creates. **[OPEN B2: any create = fail?]**
- **Todo with deadline** → `action: create_todo`, `due: { "by": "{this:FRI}" }`.
- Cardinality is **scoped to the entry's `action` + `title_match`** (e.g. "how many *board*
  events"), never a global tally. The default is exactly one, which is what makes a
  reschedule-that-left-a-duplicate fail; `count` is set only for `0` (cancel) or `N`.
- **Reply / delegate** schemas still TBD. **[OPEN]**

### 5.1 Modify / reschedule / cancel

These grade on **final state**, not on the model's mechanics — a model may reschedule by
`update_event` *or* by `delete_event` + `create_event`, and both must score identically.
Express the *intended end state*, not the edit:

```jsonc
// "Move the board meeting to the following Monday."
{ "email_id": "board-move",
  "expect": [ { "action": "create_event", "title_match": ["board"],
               "start": { "eq": "{next:MON}" } } ] }
// The default exactly-one catches the "created a new one but forgot to delete the old"
// failure — a double-booked calendar fails with no annotation.

// "Push the board meeting back a week."  (email 1 emitted {!board = next:THU@14:00})
{ "expect": [ { "action": "create_event", "title_match": ["board"],
               "start": { "eq": "@board+1w" } } ] }

// "Cancel the board meeting."   count:0 is the one cardinality you still write.
{ "expect": [ { "action": "create_event", "title_match": ["board"], "count": 0 } ] }

// "Keep the day, move 2pm → 4pm."   attach on a DATETIME overrides its time.
{ "expect": [ { "action": "create_event", "title_match": ["board"],
               "start": { "eq": "@board@16:00" } } ] }
```

---

## 6. Recurrence (start is a token; cadence is static)

```jsonc
{ "action": "create_event",
  "title_match": ["training"],
  "start": "month:+1m",          // token: which month the series begins
  "recurrence": { "every": "TUE", "at": "10:00", "dur": "60m", "count": 4 },  // all static
  "tolerance": "exact_time" }
```

Grader expands the 4 expected datetimes from the token start + static cadence and checks 4
matching events exist.

---

## 7. Cross-node templates (revised — no fact emissions)

1. **Date-on-date (the only thing that truly needs emission).**
   A: *"Signing locked for `{!signing=+5d}`."*  B: *"Kickoff two weeks after the signing."*
   → B answer: `start: eq @signing+2w`, `tolerance: exact_day`.
2. **Static fact set by an ancestor (no emission — baked in).**
   A: *"Client meetings are now 90 minutes."*  B: *"Meet `{next:THU@14:00}`?"*
   → B answer: `start: eq {next:THU@14:00}`, `duration: "90m"` (static; we know A set it; the
   A→B edge forces A served first so the model can learn it).
3. **Cadence (static) starting at a token date.**  → the recurrence block in §6.
4. **Avoidance.**  A: *"Out the `{!blackout=week_of:{+9d}}`."*  B: *"Review next week."*
   → B answer: `start: { in: "week_of:{+1w}", not_in: "@blackout" }`, `tolerance: exact_day`.

---

## 8. Grading semantics

- **Grade on cumulative final state, not the per-email diff.** A reschedule done as
  `update_event` and one done as `delete`+`create` must score identically, so the grader
  checks the calendar/todo *state* after the email, not the edit the model made. (This is
  a change from the legacy diff-based grader and is required for §5.1 to work.)
- Match a model object to an expected entry by `email_id` (model tags every create) then by
  `title_match`. Resolve date tokens against `(serve_date_of_that_email, anchor_table)` —
  **ground truth**, independent of the model's earlier actions.
- Tokens resolve to a DATE (or INTERVAL) → **day equality**; `within:Nd` relaxes the window.
- Cardinality → **exactly one by default**, scoped to the entry's action + title_match
  (over-creation, double-booking, and "forgot to cancel" all fail). Write `count` only to
  override: `count: 0` = must not exist, `count: N` = exactly N.
- Per-email result **binary** by default + a "right-action-wrong-day" diagnostic. **[OPEN B3]**

---

## 9. Authoring style guide

The grammar is exact; English isn't. Canonical phrasing, one per token, prevents drift:
`{next:THU}` ⇄ "next Thursday" (always the coming one); `{nth:3,FRI,+1m}` ⇄ "the third
Friday of next month"; `{this:FRI}` ⇄ "this Friday." A linter validates tokens and checks
every date answer traces to a token; it can't catch ambiguous prose, so the style table is
human discipline shipped with the authoring kit.

---

## 10. Open items

- **[OPEN]** Keep optional `@HH:MM` time-attach on tokens (anti-drift on the event's "when"),
  or push times to prose too for a strictly date-only grammar? (John's call.)
- **[OPEN]** Anything real emails need that a date token can't express? (multi-day events;
  fuzzy "morning/afternoon"; quarter/fiscal dates; date *ranges as the deliverable*.)
- **[OPEN B2]** No-action strictness. **[OPEN B3]** Binary vs partial-credit headline metric.
- **[OPEN]** Reply / delegate action schemas (reschedule/cancel now drafted in §5.1).
- **[DONE]** Weekday selectors relative to a *named anchor* ("the Monday after the
  signing"). Implemented as an optional `from <expr>` clause on `next:`/`this:` (e.g.
  `next:MON from @signing`, `this:FRI from @migration+1w`); with no `from`, behavior is
  unchanged (serve-relative).
- **[OPEN]** Timezone & locale lock; default valid-slot rules for `in:` answers.
- **[IMPLEMENTATION NOTE]** Grader must be **state-based** (snapshot the calendar/todos
  after each email and evaluate `expect`/`forbid`/`count` against it), not diff-based.
