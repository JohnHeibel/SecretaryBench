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
loadable example in the authoring tool ("Load the Project Atlas example"): one compact product-launch
storyline built to tour the headline constructs and show the long-horizon test **at scale**. The two
dates the saga hangs on are published in the first email and reused by later emails, so when the
runner buries the storyline under filler, those payoffs land long after their setup. That is exactly
the recall the benchmark measures. It is eight emails.

**The setup** (from VP Product) - *one event, exact time, and the anchors everything keys off*
> **Project Atlas: code freeze Monday, launch in 10 weeks.** "Code freeze is Monday 9-10 AM. Public
> launch is targeted for ten weeks later."

*Answer:* create an event **"Atlas code freeze"** at `next:MON @09:00-10:00`. Both dates are written
once in the body as **anchors**: `{!atlas_freeze = next:MON @09:00-10:00}` and
`{!atlas_launch = next:MON+10w}`. Predicate **on exactly** (`eq`) means the assistant must hit this
day and time.

**A CEO-sent email** - *the boss can be the sender*
> **Set up my launch-week press interview.** "Once Atlas launches, get my first press interview on
> the books the week after."

*Answer:* create **"Atlas press interview"** at `@atlas_launch+1w @11:00-12:00`. **From** can be the
`CEO` firing an instruction at their own assistant. It reuses `@atlas_launch`, so it is also a needle.

**A no-action FYI** - *restraint matters*
> **Re: Project Atlas trademark cleared.** "FYI only: the trademark cleared. No action needed."

*Answer:* no ops (the **"this email needs no action"** box). The assistant should create nothing;
anything it does here is a failure.

**A timed deadline** - *on or before, to the minute*
> **Atlas beta feedback: compile the results.** "The beta feedback needs to be summarized by ten
> business days after freeze, 5 PM."

*Answer:* create a **to-do** "Atlas beta feedback" due `by @atlas_freeze+10bd @17:00`. A bare `by`
deadline is day-level; a timed `by` deadline compares the object's start to the exact cutoff. Earlier
days still pass at any time.

**A needle** (from the board chair) - *connecting an earlier fact across a gap*
> **Atlas live board demo.** "The board wants a live Atlas demo two weeks after freeze, 3-4 PM."

*Answer:* create **"Atlas board demo"** at `@atlas_freeze+2w @15:00-16:00`. Because the answer reuses
an earlier date, this is a **needle**. The assistant has to recall the freeze email and compute from it.

**A reschedule** - *a conflict, authored*
> **Re: Atlas board demo, pulling it in.** "Scheduling is tight, so pull the board demo in to eleven
> days after freeze, same hour."

*Answer:* `move` **"Atlas board demo"** to `@atlas_freeze+11d @15:00-16:00`. Conflicts are always a
later email that moves or cancels an earlier thing by name, never "find an open slot." `move` inherits
the event kind from the original.

**A cancel chain** - *create, then take it back off*
> **Atlas launch-night team dinner.** "Let's celebrate with a team dinner on launch night."
> **Re: Atlas launch dinner, let's not.** "We're going to skip the dinner. Take it off."

*Answer:* first create **"Atlas launch dinner"** at `@atlas_launch @19:00-21:00`, then `cancel`
**"Atlas launch dinner"** in the later email. A cancel names the thing only, with no date or kind.

That one storyline exercises the core benchmark: exact timed events, **anchors reused across long
gaps**, a **CEO-sent email**, a **timed `by` deadline**, a **needle**, a **reschedule**, a
**cancel**, and **no-action restraint**. Load it, run the reference solver, and it scores **1.0**,
the bar every scenario you write must clear before it joins the corpus.

### Questions authors ask
- **"The answer key feels like a different tool from the email."** It's the same dates, just
  named. Write a date once in the body with the date builder (`{!name = ...}`), then reuse it
  in the answer as `@name` with the "reuse a date from an email" chips. The email and the grader
  point at the same instant by construction.
- **"What's the difference between *on exactly* and *on or before*?"** *On exactly* (`eq`) means
  the assistant must land on that exact day (and time, for an event). *On or before* (`by`) is a
  deadline: a bare date allows any day up to and including it, while a timed deadline compares the
  object's start to the cutoff.
- **"How many kinds of action are there?"** Four: create an event, create a to-do, move/reschedule,
  cancel — or tick "needs no action." That's the whole vocabulary.
- **"Writing an op per line for a busy email is tedious."** Keep the email small when you can. If
  it really asks for several actions, add one action row per thing and reuse body anchors for the dates.

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
