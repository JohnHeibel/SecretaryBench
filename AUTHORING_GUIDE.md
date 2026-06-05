# SecretaryBench Authoring Guide — Standardized Email Patterns

> **Purpose:** This guide ensures emails grade reliably while testing realistic long-horizon temporal reasoning. Follow these patterns to avoid grading surprises and support temporal randomization.

---

## 0. Core Principles (Read This First)

1. **Write for grading reliability, not cleverness.** The model's invented title must naturally contain your match keyword.
2. **All timing is relative.** Never hardcode absolute dates; use node anchors and relative offsets.
3. **Match keywords should be obvious.** If you set `match: ["HSR"]` but the email says "antitrust filing," it fails.
4. **Test at multiple temporal scales.** If your node works at 50 days, test it at 200 and 500 days—does the prose still make sense?

---

## 1. Node Structure

### What is a Node?

A **node** is one complete storyline (a deal, an acquisition, a project, a launch). It contains:
- Multiple emails with temporal dependencies
- Shared anchors (dates created and referenced within the node)
- An optional node-level anchor for temporal scaling

### Node Template

```json
{
  "id": "project_name",
  "cast": {
    "SENDER_1": "Name (Title)",
    "SENDER_2": "Name (Title)",
    "YOU": "you"
  },
  "node_anchor": "project_start",
  "emails": [ ... ]
}
```

**`node_anchor`:** The node-level reference point for all relative timing. For temporal randomization:
- `@project_start` can be serve+{50, 100, 200, 500}d
- All child emails use `@project_start + offset`
- This way, the same story works at different time scales

---

## 2. Email Types

_Tooling note: this type mix is authoring guidance only — the webapp no longer asks you to tag an email as action / FYI / junk. Whether an email needs an action is inferred from its answer key (an empty answer = no-action)._

### A. Junk Emails (5-10% of total)

**Purpose:** Add noise, test distraction/filtering.

**Pattern:**
```json
{
  "id": "node.junk_1",
  "from": "SENDER",
  "subject": "Newsletter: Q2 Market Update",
  "body": "FYI — here are this quarter's market trends... [3-4 sentences, genuinely unrelated]",
  "depends_on": [],
  "answer": { "ops": [] }
}
```

**Rules:**
- ✓ Completely unrelated to the node's story
- ✓ Obviously not actionable ("FYI", "newsletter", "reminder")
- ✓ Safe to ignore entirely
- ✗ DO NOT use as bait (don't hide an action in junk)

---

### B. No-Action Emails (10-15% of total)

**Purpose:** Test task recognition. Model sees it and correctly infers "nothing to do."

**Type B1: Plain FYI**
```json
{
  "id": "node.status_update",
  "from": "SENDER",
  "subject": "Status update",
  "body": "Quick heads up—the board approved our proposal. Just wanted you to know. No action needed from you.",
  "depends_on": [],
  "answer": { "ops": [] }
}
```

**Type B2: Informational (Sets a Fact, No Action)**
```json
{
  "id": "node.policy",
  "from": "SENDER",
  "subject": "Reminder: Office is closed week of July 4",
  "body": "The office will be closed {!blackout = week_of:(+45d)}. Just a heads-up.",
  "depends_on": [],
  "answer": { "ops": [] }
}
```
*Note: This creates an anchor but requires no action. Later emails will reference it.*

**Type B3: Implicit No-Action (Looks actionable but isn't)**
```json
{
  "id": "node.misleading",
  "from": "SENDER",
  "subject": "Do we need a kickoff meeting?",
  "body": "I'm wondering if we should schedule a kickoff. What do you think? Let me know your thoughts.",
  "depends_on": [],
  "answer": { "ops": [] }
}
```
*A T3 trap: model might think "schedule something" when the email is just asking for input.*

**Rules:**
- ✓ Genuinely require no calendar/todo action
- ✓ Can set anchors (create facts/dates)
- ✓ Can be T1 (obvious no-op) or T3 (subtle no-op)
- ✗ DO NOT: Create a no-action answer but hint at action in prose

---

### C. Action Emails (75-85% of total)

Action emails split into three tiers by difficulty:

---

## 3. Difficulty Tiers (T1, T2, T3)

_Tooling note: difficulty tiers are authoring guidance only — the webapp no longer asks you to tag T1/T2/T3._

### T1 — Easy (Self-Contained, ~30% of action emails)

**Criteria:**
- Everything needed is in *this email*
- No dependency on earlier emails
- No anchor references
- Task is explicit ("schedule X")

**Pattern:**

```json
{
  "id": "node.t1_kickoff",
  "from": "SENDER",
  "subject": "Let's schedule the project kickoff",
  "body": "Can you put the project kickoff on the calendar for next Thursday? It should be 1 hour. Thanks!",
  "depends_on": [],
  "answer": {
    "ops": [
      {
        "create": "kickoff",
        "kind": "event",
        "on": { "eq": "next:THU" },
        "match": ["kickoff"]
      }
    ]
  }
}
```

**Grading Checklist:**
- [ ] Model creates 1 event (not 0, not 2+)
- [ ] Event title contains "kickoff" (fuzzy match)
- [ ] Event is on next Thursday ✓

**Temporal Randomization:** ✓ OK
- `next:THU` works at any node scale (always "first Thursday after now")
- No issues.

---

### T2 — Medium (One Reference, Recent, ~40% of action emails)

**Criteria:**
- References ONE anchor from earlier in the same node
- Anchor is set 1-7 days back (recent, likely in context)
- Task is explicit once you recall the anchor
- Often a reschedule or follow-up

**Pattern:**

```json
{
  "id": "node.signing",
  "from": "SENDER",
  "subject": "Signing date locked",
  "body": "Great news—signing is locked for {!signing = @project_start+14d}. Put it on your calendar.",
  "depends_on": [],
  "answer": {
    "ops": [
      {
        "create": "signing",
        "kind": "event",
        "on": { "eq": "@signing" },
        "match": ["signing"]
      }
    ]
  }
}
```

Then later:

```json
{
  "id": "node.kickoff_schedule",
  "from": "SENDER",
  "subject": "Schedule the project kickoff",
  "body": "Now that signing is locked, let's schedule the project kickoff for two weeks after signing.",
  "depends_on": [
    { "email": "node.signing", "type": "date" }
  ],
  "answer": {
    "ops": [
      {
        "create": "kickoff",
        "kind": "event",
        "on": { "eq": "@signing+2w" },
        "match": ["kickoff"]
      }
    ]
  }
}
```

**Grading Checklist:**
- [ ] Depends_on links to the email that creates the anchor
- [ ] Answer key uses the anchor (e.g., `@signing+2w`)
- [ ] Match keyword is in both the email and a natural model title
- [ ] Anchor is recent enough that it's plausible the model remembers it

**Temporal Randomization:** ✓ OK
- `@signing` is created relative to `@project_start`
- `@signing+2w` is always 2 weeks after signing, regardless of node scale
- No issues.

---

### T3 — Hard (Multi-Constraint or Ambiguous, ~30% of action emails)

**Criteria:**
- Stacks 2+ hardeners:
  - **A) Long retrieval span:** Fact set 10+ days back, out of context
  - **B) Multi-constraint:** Combines conditions from 2+ earlier emails
  - **C) Task ambiguity:** Action is implied, not stated

**Pattern 1: Long Span (Hardener A)**

```json
{
  "id": "node.constraint_far_back",
  "from": "SENDER",
  "subject": "Office closure reminder",
  "body": "Just a reminder: the office will be closed {!blackout = week_of:(@project_start+40d)}.",
  "depends_on": [],
  "answer": { "ops": [] }
}
```

Then 15+ days and 10+ other emails later:

```json
{
  "id": "node.t3_multi_constraint",
  "from": "SENDER",
  "subject": "Client visit next month",
  "body": "The client wants to visit next month. Can you find a day during week 3 of the project? Preferably a Thursday.",
  "depends_on": [
    { "email": "node.constraint_far_back", "type": "static" }
  ],
  "answer": {
    "ops": [
      {
        "create": "client visit",
        "kind": "event",
        "on": {
          "in": "week_of:(@project_start+21d)",
          "not_in": "@blackout",
          "any_of": ["this:THU", "next:THU"]
        },
        "match": ["client", "visit"]
      }
    ]
  }
}
```

**Why it's hard:**
- Blackout was set 15 emails ago (must search inbox)
- Two constraints must be satisfied (week 3 AND not blackout AND Thursday)
- Model must recognize the implicit constraint (office closure applies here)

---

**Pattern 2: Task Ambiguity (Hardener C)**

```json
{
  "id": "node.t3_ambiguous",
  "from": "SENDER",
  "subject": "Legal update on HSR filing",
  "body": "Heads up: legal says the HSR filing window closes 30 days after signing, and we absolutely cannot miss this deadline. It's critical.",
  "depends_on": [
    { "email": "node.signing", "type": "date" }
  ],
  "answer": {
    "ops": [
      {
        "create": "HSR filing",
        "kind": "todo",
        "on": { "by": "@signing+30d" },
        "match": ["HSR", "filing"]
      }
    ]
  }
}
```

**Why it's hard:**
- Email does NOT say "create a task"
- Task is implied by consequence ("we cannot miss it")
- Model must infer: deadline → create a todo with that deadline
- If match keywords don't match what model invents, it fails

---

**Grading Checklist (T3):**
- [ ] Hardener A: Fact is 10+ days back (must search)
- [ ] Hardener B: Multiple constraints combined (`in` + `not_in` + `any_of`)
- [ ] Hardener C: Action is implied, not stated
- [ ] Match keywords are natural words the model would use
- [ ] Answer key is unambiguous (exactly 1 object should match)

---

## 4. Temporal Randomization Best Practices

### The Node Anchor Pattern

**All node emails should measure time from a single node-level anchor.**

```json
{
  "id": "construction",
  "node_anchor": "construction_start",
  "emails": [
    {
      "id": "construction.kickoff",
      "body": "Construction starts {!construction_start = +5d}",
      "answer": { "ops": [] }
    },
    {
      "id": "construction.framing",
      "body": "Framing phase starts {!framing = @construction_start+14d}",
      "answer": { "ops": [] }
    },
    {
      "id": "construction.roofing",
      "body": "Schedule roof inspection {!@construction_start+45d}",
      "answer": {
        "ops": [
          { "create": "roof inspection", "kind": "event", 
            "on": { "eq": "@construction_start+45d" } }
        ]
      }
    }
  ]
}
```

**Why this works:**
- `@construction_start` can be serve+{50, 100, 200, 500}d (randomized at runtime)
- All child offsets stay the same: +14d, +45d, etc.
- Same story structure, different absolute timeline
- Model tests proportional reasoning, not memorized dates

### Prose That Scales

**✓ Good (scales across 50–500 days):**
```
"Construction starts {!construction_start = +Xd}"
"Phase 2 kicks off 30 days after construction starts"
"Schedule final inspection 75 days into the project"
```

**✗ Bad (breaks at different scales):**
```
"Inspection tomorrow"            ← too rigid
"Quick check next week"          ← implies short project
"This is urgent—do it ASAP"     ← weird on 500-day timeline
"3-month project deadline"       ← hard-codes the scale
```

---

## 5. Cross-Node Dependencies

### Pattern: Node B Depends on Node A

```json
{
  "id": "moveIn",
  "cast": { ... },
  "node_anchor": "moveIn_start",
  "emails": [
    {
      "id": "moveIn.prepare",
      "from": "SENDER",
      "subject": "Ready to move in",
      "body": "Now that construction is done, let's schedule the move-in for {!moveIn_start = @construction.construction_start+90d}",
      "depends_on": [
        { "email": "construction.roofing", "type": "date" }
      ],
      "answer": {
        "ops": [
          { "create": "move-in", "kind": "event", 
            "on": { "eq": "@moveIn_start" } }
        ]
      }
    }
  ]
}
```

**Rules:**
- Email B depends on the specific email in Node A that creates the anchor
- Use `type: "date"` if it references an anchor
- Use `type: "static"` if it just needs info (no date reference)

---

## 6. Grading-Safe Patterns

### DO: Default Match Keywords to Obligation Name

```json
// ✓ GOOD: model titles it "kickoff meeting" or "project kickoff"
{ "create": "kickoff", "kind": "event", "on": { ... } }
// (match defaults to ["kickoff"])

// ✗ RISKY: model might title it "Bob's meeting" or "sync"
{ "create": "kickoff", "kind": "event", "match": ["sync"], "on": { ... } }
```

### DO: Use Obvious Keywords

```json
// ✓ GOOD: word appears in the email prose naturally
"Please schedule the HSR filing review..."
{ "create": "HSR filing", "match": ["HSR", "filing"] }

// ✗ BAD: model won't use this exact phrasing
"We need to file the antitrust notification..."
{ "create": "HSR filing", "match": ["HSR"] }  // model won't invent "HSR"
```

### DO: One Object Per Op

```json
// ✓ GOOD: answer expects exactly 1 event
{ "ops": [{ "create": "kickoff", ... }] }

// ✗ BAD: answer might create 2 kickoffs, still matches
{ "ops": [
    { "create": "kickoff", "on": { "eq": "@date1" } },
    { "create": "kickoff", "on": { "eq": "@date2" } }
] }
// Grader fails if model creates only 1, or 3
```

### DO: Use `by` for Deadlines, `eq` for Fixed Dates

```json
// ✓ GOOD: todo with deadline (on-or-before)
{ "create": "filing", "kind": "todo", "on": { "by": "@signing+30d" } }

// ✓ GOOD: event on a specific date
{ "create": "kickoff", "kind": "event", "on": { "eq": "@signing+2w" } }
```

### DO: Test Match Keywords With Fuzzy Matching

```json
// The grader matches if model's title contains ALL keywords:
"match": ["HSR", "filing"]

// These PASS:
"HSR filing reminder" ✓
"HSR Filing Review" ✓ (case-insensitive)
"Complete HSR filing" ✓

// These FAIL:
"Antitrust notification" ✗ (no "HSR")
"Filing deadline" ✗ (no "HSR")
```

---

## 7. Anti-Patterns (Don't Do This)

| **Anti-Pattern** | **Why It Fails** | **Fix** |
|---|---|---|
| `match: ["URGENT"]` | Model won't title it with all-caps | Use `match: ["urgent"]` or better: `match: ["deadline"]` |
| Multiple creates of same obligation | Grader expects exactly 1 | Split into separate scenarios or use `move`/`cancel` |
| Hard-coded absolute dates in prose | Breaks at different temporal scales | Use `{!anchor = @base+offset}` or relative language |
| Action implied but no answer ops | Creates silent no-action when meant to act | Either: keep as `ops: []` and re-write prose, OR add the ops |
| Cross-node reference without depends_on | Scheduler won't guarantee ordering | Always link with `depends_on: [{ email: "...", type: "date/static" }]` |
| Anchor with no emitter | Grader can't resolve it | Ensure ancestor email sets the anchor with `{!name = ...}` |
| Reschedule without depends_on to create | Model creates new event instead of moving | Link `move` op to the `create` op email |

---

## 8. Email Checklist Before Authoring

Use this checklist for every action email:

- [ ] **Task clarity:** Is it clear what the model should do? (T1) Or must it infer? (T3)
- [ ] **Dependencies:** Does this email reference earlier facts/dates? If yes, add `depends_on`.
- [ ] **Anchors:** Does this email CREATE an anchor (use `{!name = ...}`)? Does it REFERENCE one (use `@name`)?
- [ ] **Match keywords:** Are they words that naturally appear in the email prose?
- [ ] **Temporal scale:** Does the prose make sense at 50, 200, and 500 days? (if randomizing)
- [ ] **Grading safety:** Will exactly one object match these keywords? Or will two emails' "kickoff" events collide?
- [ ] **No-action emails:** Is this genuinely no-action, or does it hide an action?

---

## 9. Code Considerations for Implementation

### Validator Server (for Authors)

The validator should:
1. **Check anchor reachability:** Every `@name` has an ancestor that emits it
2. **Check keyword collisions:** Two obligations in the same node can't have keywords that catch each other
3. **Check temporal feasibility:** With randomized node anchors, can all emails still be served in time?
4. **Warn on risky patterns:**
   - Multiple creates of the same obligation
   - Hard-coded dates in email prose (should be `{!tokens}`)
   - Match keywords that don't appear in prose

### Grade Engine (for Results)

The grader should:
1. **Record actual anchor values:** When rendering, log what `@signing` was actually set to
2. **Resolve relative to actual values:** Don't use authored tokens; use what was rendered
3. **Report per-obligation:** PASS/FAIL per obligation, not per email
4. **Log which email created each object:** For debugging attribution

### Temporal Randomization (for Scheduler)

When serving emails:
```python
# Author specifies:
{ "id": "node", "node_anchor": "start", "emails": [...] }

# At runtime, pick one random node length:
node_length = random.choice([50, 100, 200, 500])  # seeded
node_anchor_value = serve_date + node_length

# All emails in node render relative to node_anchor_value:
actual_signing_date = node_anchor_value + 14  # if email says @start+14d
```

---

## 10. Example: Complete T3 Scenario (2 Nodes, Temporal Randomization)

```json
[
  {
    "id": "construction",
    "cast": {
      "SENDER": "Facilities Manager",
      "YOU": "you"
    },
    "node_anchor": "construction_start",
    "emails": [
      {
        "id": "construction.kickoff",
        "from": "SENDER",
        "subject": "Construction timeline",
        "body": "Construction begins {!construction_start = +7d}. Framing phase starts 2 weeks in, roofing phase 6 weeks in.",
        "depends_on": [],
        "answer": { "ops": [] }
      },
      {
        "id": "construction.blackout",
        "from": "SENDER",
        "subject": "Office will be closed",
        "body": "FYI—the office will be closed for 1 week during the roofing phase: {!roofing_week = week_of:(@construction_start+45d)}",
        "depends_on": [],
        "answer": { "ops": [] }
      },
      {
        "id": "construction.roof_inspection",
        "from": "SENDER",
        "subject": "Schedule roof inspection",
        "body": "Let's inspect the roof. How about sometime in the week after roofing finishes, but avoid the office closure week?",
        "depends_on": [
          { "email": "construction.blackout", "type": "static" }
        ],
        "answer": {
          "ops": [
            {
              "create": "roof inspection",
              "kind": "event",
              "on": {
                "in": "week_of:(@construction_start+52d)",
                "not_in": "@roofing_week"
              },
              "match": ["roof", "inspection"]
            }
          ]
        }
      }
    ]
  },
  {
    "id": "moveIn",
    "cast": {
      "SENDER": "Facilities Manager",
      "YOU": "you"
    },
    "node_anchor": "moveIn_start",
    "emails": [
      {
        "id": "moveIn.ready",
        "from": "SENDER",
        "subject": "Building ready—let's move in",
        "body": "Construction wrapped. The building is ready. Let's schedule move-in for {!moveIn_start = @construction_start+60d}. The first tenant group is flexible on which Thursday works.",
        "depends_on": [
          { "email": "construction.roof_inspection", "type": "date" }
        ],
        "answer": {
          "ops": [
            {
              "create": "move-in",
              "kind": "event",
              "on": {
                "eq": "@moveIn_start",
                "any_of": ["this:THU", "next:THU"]
              },
              "match": ["move-in", "move in"]
            }
          ]
        }
      }
    ]
  }
]
```

**Why this works:**
- ✓ Node-level anchors allow scaling (construction_start, moveIn_start randomized)
- ✓ Temporal chain (construction → roof inspection → move-in)
- ✓ T3 complexity (roofing_week constraint buried, must recall from earlier email)
- ✓ Match keywords are natural ("roof inspection", "move-in")
- ✓ Scales from 50 to 500 days without prose breaking

---

## Summary

| **Element** | **Goal** | **How** |
|---|---|---|
| **Nodes** | Organize storylines | One per event; can be parallel or dependent |
| **Junk emails** | Test filtering | 5-10%, obviously unrelated, no action |
| **No-action emails** | Test task recognition | 10-15%, can set facts but don't require action |
| **T1 emails** | Baseline accuracy | Self-contained, explicit task, no dependencies |
| **T2 emails** | Recall + accuracy | One anchor, recent, task still explicit |
| **T3 emails** | Discrimination | 2+ hardeners: span, constraints, ambiguity |
| **Temporal randomization** | Avoid time-constraint bias | Use node anchors; all offsets are relative |
| **Grading safety** | Reliable scoring | One object per op, obvious match keywords, clear dependencies |

