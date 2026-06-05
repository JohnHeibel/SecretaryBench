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

A scenario is just a short email thread with an answer for each email. Here's a full
one — "Project Helios" — to show the shape:

**Email 1** (from the COO) — *the day it arrives is the model's "now"*
> **Helios kickoff.** "We're starting Project Helios. Put the kickoff on the calendar
> for next Thursday, 10–11 AM."

*Answer:* create an event **"Helios kickoff"** at `next:THU @10:00-11:00`. Because a
later email refers back to this date, also publish it as an **anchor** so it can be
reused: `{!helios_kickoff = next:THU @10:00-11:00}`.

**Email 2** (from the COO, a week or so later) — *the long-horizon part*
> **Helios review.** "Two weeks after the kickoff, let's do a review, 2–3 PM."

*Answer:* create **"Helios review"** at `@helios_kickoff+2w @14:00-15:00`. The model
has to recall the kickoff date — possibly served weeks ago — and compute two weeks out.
That recall-and-compute over a long gap is exactly what we're testing.

**Email 3** (from the board's office) — *a reschedule conflict*
> **Re: Helios review.** "The board needs that review moved up a week, same time."

*Answer:* `move` **"Helios review"** to `@helios_kickoff+1w @14:00-15:00`. The model
must find the event it already made and move it. One exact new time.

**Email 4** (from the COO) — *a to-do with a deadline*
> **Helios board filing.** "The board filing needs to be submitted within ten business
> days of the kickoff."

*Answer:* create a **to-do** "Helios board filing" due `by @helios_kickoff+10bd`. A to-do
has no clock, and `by` means any day up to the deadline counts. It reuses the kickoff
anchor, so it is also a needle.

**Email 5** (from a teammate) — *an FYI, no action*
> "FedEx is dropping the Helios mockups Thursday, nothing for you to do."

*Answer:* no ops. The model should create nothing. This tests **restraint** (not
over-acting on chatter).

That single scenario exercises the whole benchmark: fixed times, an **anchor reused
across a long gap** (the temporal-reasoning core), a **reschedule**, a **to-do with a
due-date deadline**, and a **no-action distractor**. Run it on its own and the reference solver should score it
**1.0** before you add it to the corpus — that's your safety check.

**Want anti-gaming variety?** Write the offset as a list: `@helios_kickoff+[2,3]w`. One
run the review is 2 weeks out, another it's 3, picked by the run's seed — and the email
the model reads *and* the grader both use the same value, so it's still exactly one
right answer per run. The model can't memorize a fixed date; it has to actually reason.

### Rules of thumb for authors
- **Name it like the email calls it** ("Helios kickoff") and you're done — no keyword
  fiddling.
- **Give an exact time** for every meeting. The model must hit it exactly.
- **Use an anchor** (`{!name = ...}`) whenever a later email refers back to an earlier
  date — that's what makes it a long-horizon test instead of a one-shot.
- **Conflicts = reschedules:** a later email moves or cancels an earlier event.
- **Add FYI / no-action emails** so the model is tested on *not* over-acting.
- **Check it solves alone** before shipping it.

## Naming events (a note for authors)

Each thing the model schedules has a short name (e.g. "board meeting") so the grader can
find it on the calendar. **By default the name you give it is all you need** — you
usually don't have to think about anything fancier; the app fills it in for you. (Behind
the scenes the grader matches on those words, but that's the app's job to handle, not
yours.)
