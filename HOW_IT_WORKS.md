# How SecretaryBench Works

A plain-language tour of how the benchmark runs and how scoring works. For the
precise, technical version, see `GRADING_MODEL.md`.

## What we're testing

Can an AI assistant manage a busy CEO's calendar over a long stretch of time, working
only from the emails that land in the inbox? It has to read each email, figure out
what to schedule / move / cancel, and keep everything straight across many days.

## The simulation

- Time runs in **simulated days**.
- Each day, a **batch of emails** arrives in the inbox.
- The model reads that day's emails and uses tools (create / move / delete events and
  to-dos, search the inbox) to do what they ask.
- This repeats for many days — that long span is the "long horizon."

## Scenarios (storylines)

Emails are grouped into **scenarios** — a scenario is one storyline (e.g. "the Q3
board-meeting saga"): a handful of related emails, often spread across days, about the
same people and events. **Authors write one scenario at a time**, and each one is
self-contained and can be checked on its own.

## One calendar, many scenarios

At run time, many scenarios play out **at the same time** on **one shared calendar** —
just like a real assistant juggling everything in one place. The model sees the whole
calendar and one big inbox. It does **not** know which emails belong to which scenario;
it just sees a busy, realistic stream of work.

## Scoped grading (the key idea)

Even though everything shares one calendar, **each scenario is graded on its own.**
When we grade the board-meeting scenario, we only look at *that* scenario's events. (We
know which event came from which scenario because every event is tagged with the email
that created it.)

Why this matters: **two different scenarios can schedule something at the same time,
and that's completely fine.** If the board meeting (scenario A) and an investor call
(scenario B) both land Tuesday at 2 PM, they overlap on the calendar — but we grade A
only against A's events and B only against B's. Each one passes if it did *its own*
job. The overlap is just realistic noise; it never causes a wrong grade.

## How an email is graded

Simple and binary. For each email: **did the model do exactly what it asked, and
nothing it didn't ask?**

- Right thing, nothing extra → the email scores **1**.
- Got it wrong, or did something it wasn't asked to → scores **0**.

No half credit.

## Fixed times — no guessing

Every email says **exactly when**: "Schedule the board meeting Tuesday 2–3 PM." "Move
it to 4 PM." "Cancel it." The model just does it, and we check the calendar matches.
There is no "find an open slot yourself" — naming the time keeps every answer crisp and
gradeable. (Work hours are 5 AM–11 PM, and every time an email gives is inside that.)

The match is **exact** — the model has to land on the precise day and time, no "close
enough." (The one deliberate exception is a deadline like "file the report **by**
Friday," where any day up to Friday counts. It's rare, and you choose it on purpose.)

## Conflicts

We absolutely test the CEO's calendar getting crowded — we just **write the fix into
the story**:

> Day 3: "Board meeting Tuesday 2–3 PM." → model books it.
> Day 12: "An investor call has to take Tuesday 2–3 PM. Move the board meeting to
> 4–5 PM." → model moves it.

The model faces a real conflict and has to reschedule, but the correct outcome is fixed
(board at 4, call at 2), so it grades cleanly.

## The fence (staying in your lane)

The model should only touch events from the scenario it's currently working on. If,
while handling one email, it tries to edit or delete an event that belongs to a
**different** scenario:

- the system **refuses** the change (so the other scenario stays safe), and
- that email **loses its point.**

The model is never told this rule. A good assistant simply doesn't poke at unrelated
meetings, so a good one never trips it. We dock the point because reaching into another
scenario's stuff is a real mistake worth catching — and blocking it means one model's
slip can never damage another scenario's score.

## What the model can and can't see

- **Can see:** the inbox (all emails, searchable), the calendar (all events), its tools.
- **Cannot see:** the answer key, how it's scored, the scenario boundaries, or the
  fence. It just gets emails and does the work, like a real assistant.

## Building a scenario: a worked example

A scenario is just an email thread with an answer for each email. **"Project Atlas"** is the
loadable example in the authoring tool ("Load the Project Atlas example"): one storyline — a
company acquiring a competitor, Northwind — built to tour every kind of email *and* every feature,
and to show the long-horizon test **at scale**. The two dates the whole saga hangs on are published
in the first email and reused by emails far down the thread, so when the runner buries the storyline
under hundreds of filler emails, those payoffs land long after their setup — exactly the recall the
benchmark measures. It is 18 emails; here is a representative one for each construct.

**The setup** (from Corp Dev; **To** the CEO *and* the board, **Cc** legal + finance) — *one event, exact time, and the anchors everything keys off*
> **Project Atlas: signing Friday, target close in 12 weeks.** "Signing is locked for Friday
> 10–11 AM. We're targeting close twelve weeks out."

*Answer:* create an event **"Atlas signing"** at `next:FRI @10:00-11:00`. Both dates are written
once in the body as **anchors** — `{!atlas_signing = next:FRI @10:00-11:00}` and
`{!atlas_close = next:FRI+12w}` — so the email and the answer can never disagree. Predicate **on
exactly** (`eq`): the assistant must hit this day *and* time. (**To** / **Cc** can hold several
people and are never graded — they only make the email read true.)

**A CEO-sent email** — *the boss can be the sender*
> **Set up my 1:1 with Northwind's CEO after close.** "Once we close, get my first 1:1 on the
> books the week after."

*Answer:* create **"Northwind 1:1"** at `@atlas_close+1w @11:00-12:00`. **From** can be the `CEO`
firing an instruction at their own assistant — a normal shape, not just inbound mail. It reuses
`@atlas_close`, so it's also a needle.

**A needle** (from the board's office) — *connecting an earlier fact across a gap*
> **Atlas: board ratification vote.** "The board ratifies two weeks after signing, 3–4 PM."

*Answer:* create **"Atlas board vote"** at `@atlas_signing+2w @15:00-16:00`. "Two weeks after
signing" is just the signing anchor plus two weeks. Because the answer reuses an earlier date, this
is a **needle** — the assistant has to recall the signing (served long ago) and compute from it.
Reusing the anchor wires the dependency for you. (The headline needle is **"Atlas close ceremony"**
on `@atlas_close` itself — set twelve weeks earlier in email 1, the deepest gap in the corpus.)

**A reschedule** (from the board) — *a conflict, authored*
> **Re: Atlas board vote, pulling it in a few days.** "Quorum's tight — move the vote up, same hour."

*Answer:* `move` **"Atlas board vote"** to `@atlas_signing+11d @15:00-16:00`. Conflicts are always a
later email that moves (or cancels) an earlier thing by name — never "find an open slot." `move`
inherits the event kind from the original.

**A to-do on a deadline** (from legal) — *on exactly vs on or before, plus a title keyword*
> **Atlas: HSR filing before close.** "Our HSR filing has to be in five business days before close."

*Answer:* create a **to-do** "Atlas HSR filing" due `by @atlas_close-5bd`. A to-do has no clock and
is graded **on or before** (`by`) — any day up to and including the deadline is correct. It reuses
`@atlas_close`, so it's a long-span needle. (`match: ["HSR"]` is the rare "Advanced" override: it
tells the grader to find the task by the word *HSR*, since that's what a natural title would say.)

**Many actions in one email** (from finance) — *the scaffold button*
> **Re: Atlas diligence: three working sessions.** "Finance can do [a slot], tech wants [another],
> HR/people is [a third]. Book all three."

*Answer:* three `create` events, one per session, at `@dd_finance` / `@dd_tech` / `@dd_people`. Each
slot is written once in this email's body as an anchor and reused in the answer — so you don't
re-type three dates. In the tool, the **"scaffold an action for each date in the email"** button
builds the three rows from the body; you just name each. (An anchor a single email both defines
*and* reuses needs no dependency edge.)

**A cancel** (from comms) — *taking something back off*
> **Re: Atlas close dinner, let's not.** "We're skipping the close dinner — take it off."

*Answer:* `cancel` **"Atlas close dinner"** (created by an earlier email). A cancel names the thing
only — no date, no kind (both inherited). After this email, that event is gone.

**Any of a few days** (from people ops) — *more than one day is fine*
> **Atlas: employee town hall, pick a morning.** "Either Tuesday or Thursday next week works."

*Answer:* create **"Atlas town hall"** with **any of** `[next:TUE+1w, next:THU+1w]`. The assistant
is correct landing on *any one*. `any_of` (and the `by` deadline) are the only deliberately
multi-day answers — there is still no "find a free slot."

**The distractors** (legal, comms, people — five of them) — *FYIs, no action*
> e.g. **"Atlas: NDA countersigned"** — "No action needed, just keeping you in the loop."

*Answer:* no ops (the **"this email needs no action"** box). The assistant should create nothing;
anything it does here is a failure. This tests restraint, and a real corpus is mostly emails like
this so the action emails aren't obvious.

That one storyline exercises the whole benchmark: an exact event, **anchors reused across long gaps**
(the temporal core, stretched twelve weeks for the at-scale test), a **CEO-sent email**,
**multi-recipient To + Cc**, a **reschedule**, a **to-do on a deadline** (with a keyword override), a
**needle**, **several actions in one email**, a **cancel**, an **any-of**, and **no-action
distractors**. Load it, run the reference solver, and it scores **1.0** — the bar every scenario you
write must clear before it joins the corpus.

### Questions authors ask
- **"The answer key feels like a different tool from the email."** It's the same dates, just
  named. Write a date once in the body with the date builder (`{!name = ...}`), then reuse it
  in the answer as `@name` (the "reuse a date from an email" chips, or the scaffold button). The
  email and the grader point at the same instant by construction.
- **"What's the difference between *on exactly* and *on or before*?"** *On exactly* (`eq`) means
  the assistant must land on that exact day (and time, for an event). *On or before* (`by`) is a
  deadline: any day up to and including it counts — used for to-dos and "submit by" tasks.
- **"How many kinds of action are there?"** Four: create an event, create a to-do, move/reschedule,
  cancel — or tick "needs no action." That's the whole vocabulary.
- **"Writing an op per line for a busy email is tedious."** Put each date in the body as an anchor
  and use the scaffold button: it adds one create-event row per date with the date filled in.

### Rules of thumb for authors
- **Name it like the email calls it** ("Atlas board vote") and you're done — no keyword fiddling.
- **Give an exact time** for every meeting; use **on or before** for "submit by" to-dos.
- **Use an anchor** (`{!name = ...}`) for any date a later email refers back to — that's what
  makes it a long-horizon test instead of a one-shot — and for several dates in one busy email.
- **Conflicts = reschedules:** a later email moves or cancels an earlier thing by name.
- **Add FYI / no-action emails** so the model is tested on *not* over-acting.
- **Check it solves alone** before shipping it (`npx tsx scripts/checkTemplate.mts` does this for
  the example; the bottom-bar check does it for what you author).

## Naming events (a note for authors)

Each thing the model schedules has a short name (e.g. "board meeting") so the grader can
find it on the calendar. **By default the name you give it is all you need** — you
usually don't have to think about anything fancier; the app fills it in for you. (Behind
the scenes the grader matches on those words, but that's the app's job to handle, not
yours.)
