# Authoring a storyline, start to finish

A concrete, click-by-click walkthrough for writing one scenario without breaking it. For the
plain-language tour of the whole benchmark, read `HOW_IT_WORKS.md`; for the precise grading
contract, `GRADING_MODEL.md`. This doc is the "just show me one done correctly" version.

## Your link is your storyline

You'll be handed a link like:

```
https://secretarybench.vercel.app/?node=vp_onboarding
```

Open it and you are in **focus mode**: you see only *your* storyline, the sidebar shows only its
emails, and the validation bar at the bottom checks only your storyline. You cannot see or break
anyone else's work, and other authors editing at the same time cannot touch yours.

- **Keep the `?node=...` on the URL.** That's what scopes you to your storyline.
- The bare site (no `?node=`) is the **coordinator view** — the full list of everyone's storylines,
  where links get handed out and the final export happens. Unless that's your job, stay on your link.
- Everything **autosaves** as you type (watch the "autosaved" note in the top bar).

## The 30-second mental model

- You're writing emails that land in a busy CEO's inbox. The CEO's **AI assistant** reads each one
  and acts: it schedules / moves / cancels events and to-dos.
- For every email you also write the **answer key**: the *exact* thing a perfect assistant should do
  with it — or that it should do **nothing**.
- You **never type calendar dates.** You describe each date *relative* to when the email arrives
  ("next Monday") or to an earlier event ("two weeks after her first day"). The engine plays your
  storyline across many simulated days and fills in the real dates. The **"preview date"** box in the
  top bar only lets you *see* concrete dates while you write — it isn't saved and changes nothing.
- Grading is **all-or-nothing per email**: do exactly what's asked and nothing extra = **1 point**;
  anything off (wrong date, wrong count, touched something it shouldn't, acted on an FYI) = **0**.

## A full worked example: "VP onboarding"

Five emails about onboarding a new VP of Sales. This **exact** storyline scores **100% under the
reference solver** (verified) — so it's a safe shape to copy.

**Cast** (the people who can appear in From / To): `CEO` = you, `HR` = Dana Whitfield, `COO` =
Marcus Lee, `IT` = Sam Okafor.

---

**Email 1 — the CEO sends it** *(a note from the boss to their assistant)*
- **From:** `CEO`  **To:** `CEO`  **Subject:** Priya starts — please book her orientation
- **Body:** Heads up: Priya Rao joins us as VP Sales. Please block her first-morning orientation for `{!vp_first_day = next:MON @09:00-10:00}`.
- **Answer key:** *Create an event* named **VP orientation**, date `next:MON @09:00-10:00`.

The `{!vp_first_day = ...}` token does two jobs: it shows the date in the email **and** publishes it
as an **anchor** named `vp_first_day` so later emails can refer back to "her first day." This is a
**CEO-sent email** — perfectly normal (see below).

---

**Email 2 — the long-horizon needle** *(arrives weeks later)*
- **From:** `HR`  **To:** `CEO`  **Subject:** Onboarding review for Priya
- **Body:** Two weeks after Priya's first day, let's hold her onboarding review on `{@vp_first_day+2w @14:00-15:00}`.
- **Answer key:** *Create an event* named **onboarding review**, date `@vp_first_day+2w @14:00-15:00`.

Because the answer reuses the `@vp_first_day` anchor, this is a **needle**: the assistant has to
*recall* a date it saw weeks ago and compute two weeks out. The app auto-adds the "depends on
Email 1, with a deadline" link for you — that's the long-horizon test working.

---

**Email 3 — a reschedule (conflict)**
- **From:** `COO`  **To:** `CEO`  **Subject:** Re: onboarding review — move it up a week
- **Body:** The board wants Priya in a session that week, so move her onboarding review up by a week, same time: `{@vp_first_day+1w @14:00-15:00}`.
- **Answer key:** *Move / reschedule* **onboarding review** to `@vp_first_day+1w @14:00-15:00`.

We never ask the model to "find a free slot" — a conflict is always written into the story as an
explicit move to one exact new time.

---

**Email 4 — a to-do with a deadline** *(also a needle)*
- **From:** `HR`  **To:** `CEO`  **Subject:** Priya's compliance paperwork
- **Body:** One to-do: Priya's compliance paperwork must be filed within five business days of her start, i.e. by `{@vp_first_day+5bd}`.
- **Answer key:** *Create a to-do* named **compliance paperwork**, date `by @vp_first_day+5bd`.

A to-do has **no clock** (no `@HH:MM`). `by` means a **deadline**: landing it on any day up to and
including five business days after her start counts. `5bd` = five *business* days.

---

**Email 5 — an FYI, nothing to do** *(tests restraint)*
- **From:** `IT`  **To:** `CEO`  **Subject:** Priya's laptop & accounts ready for day one
- **Body:** Sam from IT here — Priya's laptop and accounts will be ready on her first day. Nothing for you to do.
- **Answer key:** tick **"this email needs no action."** (The assistant should create nothing.)

About 1 in 8 emails should be like this. They catch a model that over-acts on chatter.

## What the engine does with this (the "days" part)

You wrote five emails with *relative* dates and a couple of dependency links. At run time the
scheduler turns that into an actual timeline:

- It drops emails into the inbox over many simulated days (a few per day), always in dependency
  order — Email 1 before 2, 2 before 3.
- The **gap** between them is chosen by the run, not by you. So Email 1 ("first day") might land on
  simulated day 2 and Email 2 ("two weeks after") weeks later. The assistant has to remember Priya's
  first day across that gap to get the review date right. That recall-over-a-gap **is** the test.
- What you control: the **order** (depends-on) and any **deadline** (a `by` date, or a "with a
  deadline" link). What's pinned **exactly**: the answer dates, via the anchor.

So you're never really "authoring from one day." You describe the relative skeleton; the engine lays
it on a calendar, the same way every run, seeded so it's reproducible.

## Yes — the CEO can send emails

A common question: **can the CEO be the sender?** Yes. **From** can be any cast member, including
`CEO` — that's the boss firing off an instruction to their assistant (Email 1 above). It's a normal,
encouraged shape.

One thing that's always true: **From / To / Cc are presentation only and are never graded.** Who
sends an email (and who's copied) changes the realism, never the score. So a CEO-sent note and a
COO-sent note grade identically; pick whoever makes the story read true.

## The five ways people break a storyline

1. **No exact time on a meeting.** Events need a clock: `@09:00-10:00`. Only to-dos go without one.
2. **A "refers back" date typed fresh instead of reusing the `@anchor`.** If Email 2 says "two weeks
   after her first day," its answer must be `@vp_first_day+2w`, *not* a re-typed `next:MON+2w`. Reusing
   the anchor is what makes it a real long-horizon test (and keeps the two dates from drifting apart).
3. **Over-acting on an FYI.** A "nothing to do" email must create nothing — tick *"this email needs no
   action."* Don't add a calendar hold "just in case."
4. **A missing depends-on link.** When an answer reuses an earlier date, it needs a "with a deadline"
   link back to the email that published it. The app wires this automatically when you reference an
   anchor — but if the bottom bar complains about a missing date dependency, that's what it means.
5. **Times outside work hours, or backwards.** Clocks must sit inside **05:00–23:00**, and the end
   must be after the start. `@14:00-13:00` or `@04:30-05:30` is rejected.

## You're done when

The bar at the bottom reads **"Ready for export"** *and* **"oracle solves 100%."** That green means
two things: your storyline is well-formed, and a *perfect* assistant could actually carry it out
exactly as written. If it's red, it names the email to fix and why. Don't consider a storyline
finished until it's green.
