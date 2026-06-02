# SecretaryBench Authoring Guide

Welcome! This is the friendly, no-prior-experience-needed guide to writing tests for
SecretaryBench using our web app. If you can write an email and drag a block, you can do
this. You will never have to type weird date codes or touch any code.

The app is live here (no login or account needed):

**https://secretarybench.vercel.app**

Take it slow, read top to bottom the first time, and you genuinely cannot break anything —
the app refuses to let a broken test out the door.

---

## 1. What this tool is for

We are building a benchmark: a big pile of fake emails sent to an AI "secretary," to see
whether the AI does the right thing (books the right meeting on the right day, makes the
right to-do, or correctly does nothing). Your job as an author is to **write those emails
and say what the correct response is** (the "answer key"). This tool lets you do that
safely, with a live preview that uses the *real* grader, so what you see is exactly what
gets scored.

Here is the big picture. You only work in the first box. The rest is automatic:

```
   YOU ARE HERE
   ┌─────────────────────────────────────────────┐
   │  AUTHOR   →   VALIDATE   →   EXPORT / SYNC    │   →   RUN the benchmark
   │  (write    (green bar:     (download or       │       (someone runs the
   │   emails    lint + oracle   auto-pull into    │        models against
   │   + keys)   both pass)      the corpus)       │        your corpus)
   └─────────────────────────────────────────────┘
```

- **Author** — write an email, pick the date with blocks, fill in the answer key.
- **Validate** — the green/red bar at the bottom tells you if it would load and is solvable.
- **Export / sync** — when it is green, the test is ready to ship.
- **Run** — a maintainer later runs AI models against the whole collection.

You are responsible for the first two. That's it.

---

## 2. The three kinds of email (our shared vocabulary)

Every email you write is one of three kinds. We use these exact words across the whole
project, so please learn them. **All three are graded, and all three are valuable.** They
just test different things.

```
┌──────────────┬──────────────────────────────────────────────────────────────┐
│ NEEDLE       │ Answer needs a fact from an EARLIER email (an @anchor).        │
│              │ Forces the AI to REMEMBER / SEARCH BACK. The hard, juicy kind. │
│              │                                                                │
│              │ Example: Email A says "the migration is Nov 27." Much later,   │
│              │ Email B says "book the review one week after the migration."   │
│              │ To answer B, the AI must dig up the date from A.               │
├──────────────┼──────────────────────────────────────────────────────────────┤
│ TEST         │ Answer is computable from THIS email alone — no looking back.  │
│ (self-       │ Still fully graded; just no retrieval.                         │
│  contained)  │                                                                │
│              │ Example: "Let's meet next Thursday." The AI books next         │
│              │ Thursday. Everything it needs is right there.                  │
├──────────────┼──────────────────────────────────────────────────────────────┤
│ FILLER /     │ No action expected. Pure haystack noise (newsletters, FYIs,    │
│ DISTRACTOR   │ chit-chat). Graded on the AI correctly DOING NOTHING.          │
│              │                                                                │
│              │ Example: "Reminder: the gym closes at 9 tonight." The AI       │
│              │ should NOT book anything. If it over-acts, it FAILS.           │
└──────────────┴──────────────────────────────────────────────────────────────┘
```

> **Important mindset:** "not a needle" does **not** mean "not useful." A self-contained
> test checks plain competence. A filler email checks that the AI doesn't over-act on
> noise (a very common failure!). The needle is special only because its *difficulty grows*
> as we bury more filler between the two emails — that growing gap is called the **span**,
> and it is the thing the benchmark is really measuring. We need lots of all three.

---

## 3. Your first email (a guided walkthrough)

Open **https://secretarybench.vercel.app**. You'll see three regions:

```
┌───────────┬──────────────────────────────────────────────┐
│           │  HEADER: [Editor][DAG]   preview serve date   │
│  SIDEBAR  │          ┌──────────────┐   [Export corpus ⬇] │
│  (nodes   ├──────────────────────────────────────────────┤
│   and     │                                               │
│   emails) │  MAIN AREA (the editor for the selected email)│
│           │                                               │
├───────────┴──────────────────────────────────────────────┤
│  STATUS BAR (green = good, red = something to fix)         │
└───────────────────────────────────────────────────────────┘
```

### Step 3.1 — Pick or create a node

A **node** is just a folder that groups related emails (one storyline / one scenario).
The left **Sidebar** lists every node, with the number of emails in parentheses.

- To make a new one, click **+ node** at the top of the sidebar. It gets an automatic name
  like `node-1`.
- Hover a node to reveal a small **✕** to delete it.

### Step 3.2 — Set up the cast (who's in the scenario)

Click a node and, in the main area, you'll see a collapsible **Cast** section at the top
(click the header to open it). The cast is the list of people in this storyline. Each
person has:

- a short **key** (left box, like `CEO` or `PERSON_2`) — this is the internal label, and
- a **display name** (right box, like `Dana Lee (CEO)`).

Click **+ person** to add someone. Every node starts with `CEO → you`, meaning the AI
secretary is acting on *your* behalf. Add as many people as your scenario needs.

### Step 3.3 — Add an email

Under the node in the sidebar, click **+ email**. It appears with an automatic id like
`node-1.e1`. Click it to open it in the editor.

> **Ids are automatic and you don't type them.** They're simple slugs like `node-1.e1`.
> Just leave them alone.

### Step 3.4 — Fill in From / To / Subject

In the editor you'll see:

- **From** — a dropdown of your cast keys. Pick who sent the email.
- **To** — a dropdown of your cast keys. Pick the recipient. (One recipient — see gotchas.)
- **Subject** — a normal text box. This is also what shows in the sidebar, so give it a
  clear name.

### Step 3.5 — Write the body

The **Body** is a normal text box — write the email like a human would. The one rule:
**any date must be inserted with the block builder, never typed by hand.** That's the next
section, and it's the heart of the whole tool.

Everything saves automatically as you type. There is no "save" button.

---

## 4. The date-token block builder

### Why you never type a raw date

If you typed "next Thursday" as plain text, the grader couldn't line it up with your answer
key, and the two could silently drift apart (this was a real bug we are killing). Instead,
dates are inserted as little **tokens** built from drag-and-drop blocks — the same blocks
build the email body *and* the answer key, so they can never disagree.

### Opening the builder

In the Body section, click the green **+ insert date token** button. A big dark window pops
up titled **Date token builder**. (In the answer key, the equivalent button is **build
date** — same window, same blocks.)

If you've never used Blockly (the thing Scratch is built on): it's a visual editor where
you **drag puzzle-piece blocks from the left strip** into the open canvas and **snap them
together**. Blocks with a matching notch click into each other's empty slots.

```
   left strip (palette)        canvas (drag here & snap)
   ┌─────────────┐             ┌───────────────────────────────┐
   │ serve       │   drag →    │   ┌──────────┐                 │
   │ anchor @    │             │   │ next FRI │  ← a finished   │
   │ ± offset    │             │   └──────────┘    date         │
   │ next ·      │             │                                 │
   │ this ·      │             └───────────────────────────────┘
   │ ...         │
   └─────────────┘
```

Use the trash can or the **zoom / pan controls** in the corner if things get crowded.

### The live preview (your safety net)

At the bottom of the builder there's an **Expression** line and a live result:

- **green `→ Sunday, June 14, 2026`** = it resolves correctly. 
- **red `→ ...error`** = the build is incomplete or invalid; keep going.
- **amber `→ uses an anchor; resolves at serve time`** = it depends on a date published
  elsewhere, so it can only be pinned down when the test actually runs. That's normal and
  fine for needles.

This preview runs through the **real grader**, so what you see really is what gets scored.
The **preview serve date** in the header (default `2026-06-01`) is the pretend "today" used
for the preview — change it to sanity-check how "next Thursday" lands on different days.

When you're happy, click **Insert** and the token drops into your text at the cursor. Back
in the body you'll see a little chip like `{next:THU} → Thursday, June 4, 2026` confirming it.

### What each block means

You build dates by snapping these together. Most dates start with **next**, **this**, or
**serve**.

| Block | Plain English |
|-------|---------------|
| **serve (today)** | The day the email is delivered to the secretary. Your starting "now." |
| **anchor @ ___** | A date that *another* email published. Pick its name from the dropdown. This is how a needle reaches back to an earlier fact. (Empty until some email publishes an anchor — see 4.x below.) |
| **± offset** (`__ + N units`) | Shift a date forward/back by N **calendar days / business days / weeks / months / years**. Plug another date into its left slot. E.g. "serve + 5 calendar days." |
| **next ___** (weekday) | The next given weekday strictly after today (or after an optional "from" date you plug in). "next FRI." |
| **this ___** (weekday) | The given weekday inside *this* Mon–Sun week (of today, or of a "from" date). "this WED." |
| **the Nth ___ of ___** | The Nth weekday of a month, e.g. "the 3rd Friday of next month." N can be 1–5 or **last**. Plug a month block into the "of" slot. |
| **day N of ___** | A specific day-of-month, e.g. "day 15 of this month." Plug a month block in. |
| **the week of ___** | A whole Mon–Sun **week** (an interval, not a single day) containing the date you plug in. |
| **the whole month ___** | A whole **month** as an interval. Plug a month block in. |
| **___ month** (the month picker) | Says *which* month, relative to today: **this month**, **N months ahead +**, or **N months back −**. This block feeds the "of ___" slots above; it isn't a date by itself. |

You can nest these. "The 1st Monday of 2 months ahead, plus 3 business days" is just a few
blocks snapped together — and the preview shows you the real result as you go.

### Publishing an anchor (the SETUP half of a needle)

At the bottom of the builder, when you're inserting into a **body**, there's a checkbox:

> ☐ **publish this date as an anchor named** `________`

Tick it and type a short name (letters/underscores, like `signing` or `migration`). This
does two things:

1. The date still appears in the email text as normal.
2. It **registers that date under your chosen name**, so a *later* email can reach back and
   say "@signing" in its answer key.

**This is exactly how you create the first half of a needle.** Email A publishes
`@migration`; Email B's answer key then refers to `@migration`. (Full recipe in section 6.)

> Note: a token that publishes an anchor shows in your body as `{!migration = ...}`. You
> don't type that — the checkbox writes it for you.

---

## 5. The answer key (what the right response is)

Below the body and dependencies is the **Answer key** box. This is where you declare what a
perfect secretary should do with this email. Every email needs an answer key — even filler.

### "No action expected" = this is filler

At the top right of the Answer key box is a checkbox:

> ☐ **no action expected (FYI / distractor)**

Tick it for a **filler / distractor** email. The email is then graded as **do nothing** —
if the AI creates any event or to-do for it, that's a failure (over-acting). That's the
whole test. Done.

### Adding an expected action (tests and needles)

Leave that box unticked and click **+ expected action**. You get a row with:

- **Action type** dropdown — what the secretary should do:
  - `create_event` — put a meeting on the calendar (has a start date)
  - `create_todo` — add a to-do (has a due date)
  - `reschedule` — move an existing event (has a date)
  - `reply` — send a reply (no date)
  - `delegate` — hand it to someone (no date)
- **title keywords** — comma-separated words that should appear in the title (e.g.
  `kickoff, planning`). The grader checks these loosely, so a couple of distinctive words
  is plenty.
- **count** — how many of this action are expected (almost always `1`).
- **tolerance** — how exact the date must be. `exact_day` means the day must match; you can
  also allow slack like `within:2d`.

### The date predicate (for actions that have a date)

For `create_event`, `create_todo`, and `reschedule`, a date row appears. You choose an
**operator** and then click **build date** (the same block builder from section 4) to fill
in the date. The operators:

| Operator | Meaning |
|----------|---------|
| **on exactly (eq)** | The action must land on this exact date. |
| **on or before (by)** | A deadline — any day up to and including this date is correct. |
| **within interval (in)** | Must fall inside an interval (use a "week of" / "whole month" block). |
| **any of (any_of)** | Several acceptable dates; any one counts. Build each date and they accumulate. |
| **not within (not_in)** | Must NOT fall inside the given interval. |

### The crucial part — what makes an email a needle

When you build that date and it uses an **`@anchor`** block (a date published by *another*
email), **that is what turns this email into a NEEDLE.** The answer can't be computed from
this email alone — the AI has to retrieve the earlier fact. If instead the date is built
only from `serve` / `next` / `this` / etc., the email is a self-contained **test**.

So: same answer-key UI, and the single choice of "did I plug in an `@anchor`?" decides
needle vs. self-contained test.

---

## 6. How to author a NEEDLE, end to end (the recipe)

A needle is **two emails plus a link between them.** Here's the whole thing:

```
   EMAIL A (the SETUP)                       EMAIL B (the PAYOFF)
   ────────────────────                      ────────────────────
   Body mentions a date and                  Answer key books something
   PUBLISHES it as an anchor.                 RELATIVE to that anchor.
                                              
   "Migration is {!migration = ...}"   ───►   create_event, date = @migration + 1 week
        (tick "publish as anchor",            (build date → use the "anchor @ migration"
         name it: migration)                   block, plus an offset block)
                                              
                          + a "date" dependency edge A → B
```

Step by step:

1. **Write Email A.** In its body, insert a date token and tick **publish this date as an
   anchor named** → type `migration` (or whatever fits). Insert. Email A is now the setup.
   (Email A itself can be filler — often the setup fact arrives in an otherwise no-action
   email.)
2. **Write Email B**, somewhere later in the storyline.
3. In **Email B's answer key**, add an expected action, click **build date**, and use the
   **anchor @** block (pick `migration` from its dropdown). Snap on an **± offset** block if
   the payoff is "one week after," etc. The preview will say *uses an anchor; resolves at
   serve time* — that's correct.
4. **Add a dependency edge** from A to B and set it to **date** (next section). This tells
   the system Email A must arrive before B, and that the link carries a deadline.

That's a needle. When the benchmark later buries lots of filler between A and B, the
**span** grows and the test gets harder — exactly what we want.

> **Anchor names must be unique** across the corpus, and the **@ dropdown only lists
> anchors that actually exist** — so if you don't see your anchor in Email B, double-check
> that Email A really published it (look for the `emits @name` note on Email A in the DAG
> view).

---

## 7. Depends-on edges (saying which email comes first)

Below the body is a **Depends on** box. Use it to say "this email needs an earlier one
first." Pick a prerequisite email from the dropdown, then click one of:

- **+ static** — a **fact / setup** link, *no deadline attached*. Use this when the earlier
  email just provides context. In the DAG these are the **amber** edges, and the chain of
  them is the **retrieval span** (how far back the AI has to look).
- **+ date** — a link that **carries a deadline**. Use this for the A→B link of a needle,
  or any time the later action's timing depends on the earlier email. These are the
  **blue** edges in the DAG.

Each listed dependency has a dropdown to switch its type later, and a **✕** to remove it.

> Rule of thumb: if the earlier email sets a date your answer key reaches back to, use
> **date**. If it's just background, use **static**.

---

## 8. The DAG view (seeing the structure)

Click **DAG** in the header to switch from the editor to a map of your whole corpus. ("DAG"
just means a diagram of boxes-and-arrows with no loops.) Each email is a box; arrows show
dependencies. Prerequisites sit to the **left** of the emails that depend on them.

The legend across the top tells you everything:

- **amber lines** = **static** edges (the retrieval span).
- **blue lines** = **date** edges (a carried deadline); these are animated.
- a violet **`emits @name`** note on a box = that email **publishes** an anchor (a setup).
- a grey **`no-action`** note = that email is **filler**.

Click any box to jump straight back to editing that email. The DAG is the fastest way to
eyeball "is my needle wired up?" — you should see the setup box with `emits @migration` and
a blue arrow running to the payoff box.

---

## 9. The status bar (green = good, red = fix me)

The bar at the very bottom is your single source of truth. It runs the **real benchmark
checks** automatically as you work — you don't press anything.

It goes through two gates:

1. **Lint** — would the benchmark even load this corpus? Checks the structure, ids,
   grammar, anchor references, no cycles, etc.
   - Red: **✗ won't pass the benchmark loader** + the specific problem. Fix what it says.
2. **Oracle** — once lint is green, a *perfect reference secretary* tries to solve every
   answer key. This catches answer keys that are technically valid but **impossible to
   satisfy**.
   - Red: **✗ unsatisfiable answer key** and it lists the email ids that no model could
     solve. Usually means a date predicate contradicts itself or an anchor math is off.

When everything is good you'll see the happy green message:

> **✓ corpus valid** — N nodes · N emails · N anchors · **oracle solves 100%** — ready to
> export & run

**You cannot export a broken corpus.** If the bar isn't green, the test isn't ready, and
that's the app protecting you. Green means a real model run could legitimately use it.

---

## 10. Export / how it reaches the benchmark

Once the bar is green:

- Click **Export corpus ⬇** in the header. This downloads the corpus as the exact
  `nodes/*.json` files the benchmark reads — nothing is reformatted or invented. A
  maintainer drops these into the `corpus/` folder and runs the models.

There's also an **automated pull**: maintainers can run `python -m sb.sync` to pull the
latest authored corpus straight from the app's database into the repo, so in practice you
often don't even need to hand off a file — just get to green and let the sync grab it.

Either way, your part ends at the green bar.

---

## 11. Common mistakes & gotchas

A short list of things that trip people up. None of them are dangerous — the validator
catches almost everything — but knowing them saves time.

- **Over-acting distractors.** The most common real-world AI failure is acting on filler.
  So please write plenty of tempting-but-no-action filler (an email that *sounds* like it
  wants a meeting but doesn't). Mark it **no action expected** and it's graded on the AI
  staying still.
- **Don't type dates by hand.** Always use the block builder, in both the body and the
  answer key. Typed-in dates won't be graded and defeat the whole point.
- **Ids are simple slugs and automatic.** Don't rename them to fancy strings; let the app
  assign `node-1.e1` style ids.
- **Editing a token reopens the builder fresh.** The builder doesn't reload your previous
  blocks — to change a date, open the builder, rebuild the date, and insert again (then
  remove the old token text if needed). Quick, but don't expect it to remember.
- **Single recipient.** The **To** field is one person. Don't try to address several people
  in one email.
- **Anchor names must be unique** across the whole corpus, and a payoff can only reference
  an anchor that some earlier email actually published. If the **@** dropdown is empty or
  missing your name, go publish it first (and check the DAG for the `emits @name` note).
- **Anchor dates preview as amber** ("resolves at serve time"). That is expected and
  correct — an anchor can't show a concrete date until the run pins it down. It is *not* an
  error.
- **Watch the green bar before you walk away.** If it's red, your test isn't shipped yet.
  Read the message; it tells you exactly which email and what's wrong.

---

That's everything. Make a node, write a couple of self-contained tests to warm up, sprinkle
in some tempting filler, then try a full needle with an anchor. Watch the bar go green, hit
Export (or let sync grab it), and you've contributed a real benchmark test. Thank you —
have fun!
