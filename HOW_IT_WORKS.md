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

A scenario is just a short email thread with an answer for each email. "Project Helios"
is the loadable example in the authoring tool ("Load the Project Helios example"); it is
one storyline that tours every kind of email you'll ever write, in the order you tend to
meet them. Here is the whole thing, with the question each email answers:

**Email 1** (from the COO) — *the basics: one event, one exact time*
> **Helios kickoff.** "We're starting Project Helios. Put the kickoff on the calendar
> for next Thursday, 10–11 AM."

*Answer:* create an event **"Helios kickoff"** at `next:THU @10:00-11:00`. The date is
written once in the body as an **anchor** — `{!helios_kickoff = next:THU @10:00-11:00}` —
and the answer reuses the same expression, so the email and the answer can never disagree.
Predicate **on exactly** (`eq`): the assistant must hit this day *and* time.

**Email 2** (from the COO, a week or so later) — *connecting an earlier fact (a needle)*
> **Helios review.** "Two weeks after the kickoff, let's do a review, 2–3 PM."

*Answer:* create **"Helios review"** at `@helios_kickoff+2w @14:00-15:00`. "Two weeks after
the kickoff" is just the kickoff anchor plus two weeks. Because the answer reuses an earlier
date, this is a **needle** — the assistant has to recall the kickoff (served weeks ago) and
compute from it. Reusing the anchor wires the dependency for you.

**Email 3** (from the board's office) — *a reschedule conflict*
> **Re: Helios review.** "The board needs that review moved up a week, same time."

*Answer:* `move` **"Helios review"** to `@helios_kickoff+1w @14:00-15:00`. Conflicts are
always authored as a later email that moves (or cancels) an earlier thing by name — never
"find an open slot." `move` inherits the event kind from the original.

**Email 4** (from the COO) — *a to-do with a deadline (on exactly vs on or before)*
> **Helios board filing.** "The board filing needs to be submitted within ten business
> days of the kickoff."

*Answer:* create a **to-do** "Helios board filing" due `by @helios_kickoff+10bd`. Two things
differ from an event: a to-do has no clock, and it's graded **on or before** (`by`) instead
of **on exactly** (`eq`) — landing on *any* day up to and including the deadline is correct.
(`+10bd` = ten *business* days.) It reuses the kickoff anchor, so it's also a needle.

**Email 5** (from BizDev) — *many actions in one email*
> **Three partner intros to schedule.** "Acme can do [a slot], Globex prefers [another],
> Initech offered [a third]. Please get all three on the calendar."

*Answer:* three `create` events, one per partner, at `@acme_slot` / `@globex_slot` /
`@initech_slot`. Each slot is written once in this email's body as an anchor and reused in
the answer — so you don't re-type three dates. In the tool, the **"scaffold an action for
each date in the email"** button builds these three rows for you from the body; you just
name each. (An anchor a single email both defines *and* reuses needs no dependency edge.)

**Email 6** (from BizDev) — *taking something back off*
> **Re: Globex meeting cancelled.** "Globex had to back out. Take it off the calendar."

*Answer:* `cancel` **"Meet with Globex"**. A cancel names the thing only — no date, no kind
(both inherited). After this email, that one meeting is gone and the other two remain.

**Email 7** (from the COO) — *more than one day is fine*
> **Helios offsite, pick a day.** "Either next Monday or next Wednesday works — put it on
> whichever."

*Answer:* create **"Helios offsite"** with **any of** `[next:MON, next:WED]`. The assistant
is correct if it lands on *any one* of the listed days. `any_of` (and the `by` deadline) are
the only deliberately multi-day answers — there is still no "find a free slot."

**Email 8** (from Design) — *an FYI, no action*
> "FedEx is dropping the Helios mockups Thursday, nothing for you to do."

*Answer:* no ops (the **"this email needs no action"** box). The assistant should create
nothing; anything it does here is a failure. This tests restraint, and a real corpus is
mostly emails like this so the action emails aren't obvious.

That one storyline exercises the whole benchmark: an exact event, an **anchor reused across a
long gap** (the temporal core), a **reschedule**, a **to-do on a deadline**, **several actions
in one email**, a **cancel**, an **any-of**, and a **no-action distractor**. Load it, run the
reference solver, and it scores **1.0** — which is the bar every scenario you write must clear
before it joins the corpus.

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
- **Name it like the email calls it** ("Helios kickoff") and you're done — no keyword fiddling.
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
