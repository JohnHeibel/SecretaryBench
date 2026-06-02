# Recap — day loop, grading, and what's safe to author (2026-06-02)

Read this first in the morning. It explains the whole machine in plain terms, what
changed last session, and the short list of things to keep in mind before you start
feeding real emails. Verdict up front: **the machinery is solid. You can start
authoring.** The only real risks left are authoring-discipline ones, and most are now
blocked by the linter automatically.

---

## 1. The whole pipeline in one picture

```
  corpus/nodes/*.json          sb/scheduler.py            sb/live/runner.py            sb/grader.py
  ┌──────────────────┐         ┌───────────────┐          ┌──────────────────┐         ┌────────────┐
  │ you author here  │  ────▶  │ DAG -> a serve │  ─────▶  │ per-DAY loop:    │  ────▶  │ state-based│
  │ (emails + answer │         │ plan (which    │          │ model pulls its  │         │ grade per  │
  │  keys, tokens)   │         │ email on which │          │ own inbox & acts │         │ obligation │
  └──────────────────┘         │ day)           │          └──────────────────┘         └────────────┘
        ▲                      └───────────────┘                                              │
        │                                                                                     ▼
   webapp (lint gate)                                                                   PASS / FAIL
   validates with the SAME sb code before you can save
```

The key idea: **you author emails grouped into nodes. The scheduler turns the whole
set into a day-by-day serve plan. The model handles one day at a time by pulling its
own inbox. The grader checks the calendar state after each day.** The webapp is just a
safe authoring front-end that validates against the exact same `sb` code the grader
uses.

---

## 2. What a "node" is, and how a day gets built

A **node** is one storyline. `henderson` (a deal), `globex-acq` (an acquisition), etc.
Emails inside a node can depend on each other (signing -> kickoff -> move the kickoff).

```
  NODES (storylines)                       A DAY'S INBOX (what the model sees)
  ┌─────────────┐ ┌─────────────┐          ┌─────────────────────────────────┐
  │ henderson   │ │ globex-acq  │          │ Tue Jun 02 — 3 new emails:      │
  │  intro      │ │  diligence  │  ──mix──▶ │   • henderson.kickoff (henders) │
  │  signing    │ │  kickoff    │          │   • globex.diligence  (globex)  │
  │  kickoff    │ │  filing     │          │   • pr.release        (pr-comms)│
  └─────────────┘ └─────────────┘          └─────────────────────────────────┘
```

How the scheduler picks each day (seeded, deterministic):

- It serves **every email exactly once** over the run. It does NOT take a random
  subset. (If you ever want sampling, that's a feature to add later.)
- An email is only eligible on a day if **all its dependencies were served on an
  earlier day**. So a dependent email never lands the same day as the email it needs.
- Deadline-bound emails get forced in first (so nothing misses its date window).
- The rest is shuffled to hit a per-day target size (currently 1 to 5 emails/day).

Same `(corpus, start_date, seed)` always produces the same plan. 100% reproducible.

---

## 3. The day loop (this is what we rebuilt last session)

**Before:** one model turn per email, with the email body pasted into the prompt. The
model never triaged, never discovered its own mail. That understated the task.

**Now:** one model turn per **day**, and the prompt is blind.

```
  Day begins
     │  runner tells the store "today = Jun 02"
     │  runner drops the day's emails into the store inbox (bodies live here, NOT the prompt)
     ▼
  ┌──────────────── ONE model turn ────────────────┐
  │ prompt: "Today is Jun 02. New mail has          │
  │          arrived. List your inbox, work         │
  │          through it, then stop."                │
  │                                                 │
  │  model: list_new_emails()  -> [ids+subjects]    │   ← it must DISCOVER the mail
  │         get_email(id)      -> full body         │   ← reads each one
  │         create_event / create_todo / ...        │   ← acts, stamping email_id
  │         (or does nothing for an FYI)            │
  └─────────────────────────────────────────────────┘
     │  runner snapshots calendar/todo state
     ▼  grades each email in the day's batch
```

Why this is the right call for **long-horizon temporal reasoning**: it's all one
continuous context across many days. The model carries anchors set weeks ago, and on a
long run it even has to deal with **context compaction**, exactly like a real assistant
whose memory gets summarized. A per-email loop couldn't test any of that.

---

## 4. How grading actually works (the part that surprised us)

Grading is **state-based**: after the day, look at the calendar/todos and reconcile each
obligation. It is NOT "did the model emit the right token." The model never sees or
emits tokens at all. Tokens are an authoring + grading device; the model just gets plain
English and makes plain calendar events with titles it invents.

So grading is two steps:

```
  answer key:  { "create": "kickoff", "kind": "event",
                 "on": { "eq": "@signing+2w" }, "match": ["kickoff"] }

  STEP 1 — FIND the object          STEP 2 — CHECK its date
  ───────────────────────          ───────────────────────
  among events in this node,        resolve @signing+2w from
  keep those whose TITLE            ground truth -> Jun 21,
  contains the match keyword        require exactly ONE found
  "kickoff"                         object, on Jun 21
        │                                   │
        ▼                                   ▼
  "Henderson project kickoff" ✓       Jun 21 == Jun 21 ✓  -> PASS
```

- **The token (`on`) is the real grade** — the date is the substance.
- **`match` keywords are just the finder** — how the grader picks the one object out of
  a full calendar. It defaults to the obligation's name, fuzzy substring, case-
  insensitive, checks title + description. You usually write nothing.
- **Exactly one.** A reschedule that leaves a stale duplicate fails (two matches). A
  cancel must leave zero.
- **`ops: []`** = a no-action / FYI / bait email. The turn must create nothing.

### Attribution is node-level, not email-level

When the model creates an object it stamps an `email_id`. The grader uses that **only to
route the object to its node**, then matches by title within the node. So stamping a
*sibling in the same node* is harmless; only a *wrong-node* stamp misfiles it. Since
nodes are distinct storylines, a capable model won't cross them. Mixed-node days are
intended and safe.

---

## 5. The answer-key grammar, in brief

An email's `answer` is a list of `ops`, each a verb on a named obligation:

```jsonc
{ "ops": [
  { "create": "kickoff", "kind": "event", "on": { "eq": "@signing+2w" } },
  { "move":   "kickoff", "on": { "eq": "@kickoff+3bd" } },   // later email, same node
  { "cancel": "kickoff" }                                    // no on, no kind
] }
```

- **Verbs:** `create` | `move` | `cancel`. The value is the obligation name.
- **`kind`:** `event` | `todo` (create only).
- **`on` predicate:** `eq` (exact day), `by` (deadline, on-or-before), `in` (within an
  interval), `in` + `not_in` (within a window, avoiding a blackout), `any_of` (any of a
  list).
- **Date tokens:** `serve+2w`, `+3bd` (business days), `next:THU`, `nth:3,FRI,+1m`,
  `@signing+2w`. They resolve to **whole days** (no times — times are prose, never graded).
- **Cross-email link:** the only dynamic dependency is **date-on-date**. An ancestor
  emits `{!signing=+5d}`, a descendant references `@signing+2w`. Everything else (names,
  kinds, keywords) is static.

Full reference: `ANSWER_KEY_GRAMMAR.md`.

---

## 6. What changed last session (the diff)

1. **Day loop rewrite** (`sb/live/runner.py`, `store_app.py`, `mcp_app.py`): one turn per
   day, blind prompt, model pulls its own inbox via a new `list_new_emails()` tool. Store
   gained a "today" notion (`/day`, `/inbox/new`). Grading splits the day's new objects
   back per email by the stamped `email_id`. Verified live on Haiku: 13/13.
2. **Display fix**: the `tools` / model-narration line is now printed once per day, not
   per email (it was falsely making no-action emails look like they acted).
3. **Node-name uniqueness test**: locked in (was already enforced, now has a test).
4. **Match-keyword collision lint** (`sb/schema.py`, check #5): within a node, two
   same-kind obligations can't have keyword filters that catch each other. Substring-
   aware. Re-vendored into the webapp, so the **save gate now blocks it**. Full suite 51/51.

---

## 7. What to worry about before feeding emails

Short version: **not much, and the linter now catches most of it for you.** The webapp
won't let you save a corpus that:

- has a dependency cycle,
- references an anchor with no emitting ancestor,
- uses an anchor in an answer without a date edge,
- has an unparseable token,
- duplicates a node id or email id,
- has colliding match keywords in a node. ✅ (new)

The two things the linter **cannot** check, so they're on you:



2. **Lock the run config for comparability.** `seed` and `start` silently change *what*
   gets tested (which order, which weekdays the selectors resolve to). The difficulty
   levers (`daily_min`/`daily_max` pile size, `urgency_horizon`) are currently hardcoded
   defaults and not recorded with results. For a real benchmark, pin one config (fixed
   seed, fixed start, fixed levers) and stamp it into the output so a score means the
   same thing across models.

One operational note: a day-turn does more work than an email-turn, so raise
`per_turn_timeout` for long runs or big batches. And on a new model's first run, log the
raw `email_id` stamps once so you can tell a real reasoning miss from a stamping slip.

The thing that actually matters for the benchmark's value is not covered by any lint:
**write discriminating emails.** A "discriminating" email is one that good models pass
and weak models fail. An email everyone gets right (or everyone gets wrong) is dead
weight. Aim for real cross-day spans, references that need recall, and distractors that
look actionable but aren't.

---

## 8. Next steps

- [ ] **Finalize the webapp** so you can import emails (your stated next task). The new
      collision check is already wired into its lint gate, so authoring there is safe.
- [ ] Start authoring real, human-written emails. Keep AI filler "do-nothing" emails on
      the back burner.
- [ ] Decide and pin the locked run config (seed / start / levers) before the first
      comparison run.
- [ ] Optional: add `email_id`-stamp logging to the runner for first-run diagnostics.

Files touched this session: `sb/live/runner.py`, `sb/live/store_app.py`,
`sb/live/mcp_app.py`, `sb/schema.py`, `sb/tests/test_schema.py`, and the re-vendored
`webapp/api/_lib/sb/*` + `webapp/seed/nodes.json`. Not committed yet.

--- 
## 9

- things that just don't make sense to me right now. 
1. **Pick realistic match keywords.** The grader matches the *title the model invents*.
   If you set `match: ["HSR"]` but a reasonable assistant titles it "File antitrust
   notification," it fails despite being correct. Use a word the email itself leans on
   and any sane title would echo. Prefer one distinctive noun over a multi-word set
   (it's ALL keywords). When in doubt, lean on the default (`match` = the obligation
   name) and name the obligation something a title would naturally contain. 
^ how is the grader supposed to check against tile the model invents. that can not be deterministic that makes no sense to me. 

- We need to make a tier list of how difficult the emails we will feed the model and what it should see. 
Like t1 - t3 being easy, medium, and hard should divy up the work for this. 