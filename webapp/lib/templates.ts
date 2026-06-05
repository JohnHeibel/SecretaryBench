// Ready-made scenarios authors can load as a starting point. The flagship is "Project Helios", the
// worked example in HOW_IT_WORKS.md. It is a single storyline (one cast, one DAG) deliberately built
// to exercise EVERY construct an author will ever reach for, in the order you tend to meet them, so a
// new author can load it and read one finished thread instead of guessing. It lints clean AND solves
// 1.0 under the reference oracle (verified — see scripts/checkTemplate.mts).
//
// The eight emails, and the question each one answers:
//   1. kickoff       create an EVENT at an exact time            "how do I just put one thing on the calendar?"
//   2. review        reuse an earlier date (+2w) = a NEEDLE       "how do I say 'two weeks after the kickoff'?"
//   3. review-moved  MOVE / reschedule that event                "a later email changes an earlier plan — now what?"
//   4. filing        a TO-DO graded ON OR BEFORE a deadline       "what's the difference between on-exactly and a deadline?"
//   5. partners      MANY actions in one email (the scaffold)     "writing an op per line is tedious — can I batch it?"
//   6. globex-cancel CANCEL one of those                          "how do I take something back off?"
//   7. offsite       ANY OF a few acceptable days                 "what if either day is fine?"
//   8. mockups-fyi   NO action (a distractor)                     "what about an FYI the assistant should ignore?"
//
// Two rules worth internalizing from this example:
//   - The body and the answer point at the SAME date. A date written once as {!name = expr} in a body is
//     reused in the answer as @name, so they can never disagree. (That's why the answer key isn't "a
//     different tool" from the email — it's the same dates, named.)
//   - "comes after" vs "with a deadline" edges: a reschedule/needle that reuses an earlier DATE carries a
//     `date` edge (the grader derives a serve-by window from it). Plain ordering is `static`. The app wires
//     the `date` edge for you whenever an answer reuses an @anchor; you rarely set it by hand.
import type { CorpusNode } from "./types";

// `nodeId` lets the caller pick a non-colliding storyline id; email ids are prefixed with it so the
// dependency edges (and the eventual corpus/nodes/<id>.json) stay readable.
export function projectHeliosNode(nodeId = "project_helios"): CorpusNode {
  const e = (slug: string) => `${nodeId}.${slug}`;
  return {
    id: nodeId,
    cast: { CEO: "you", COO: "Dana Reyes (COO)", BOARD: "Board office", DESIGN: "Priya Nair (Design)", BIZDEV: "Marcus Lee (BizDev)" },
    emails: [
      {
        // 1. THE BASICS. One event, one exact time. The date is written once in the body as a named anchor
        // {!helios_kickoff = ...} and the answer reuses the same expression, so they resolve to one instant.
        // Predicate is "on exactly" (eq): the assistant must land on this day AND time, to the minute.
        id: e("kickoff"), from: "COO", to: "CEO", subject: "Helios kickoff",
        body: "Hi, we're kicking off Project Helios. Please put the kickoff on the calendar for {!helios_kickoff = next:THU @10:00-11:00}. Thanks!",
        depends_on: [],
        answer: { ops: [{ create: "Helios kickoff", kind: "event", on: { eq: "next:THU @10:00-11:00" } }] },
      },
      {
        // 2. CONNECTING EARLIER FACTS (a needle). "Two weeks after the kickoff" = @helios_kickoff+2w. Because
        // the answer reuses the kickoff's date, this is a needle: the assistant has to find the kickoff email
        // to answer this one. Reusing an @anchor auto-adds the `date` edge below (you don't wire it by hand).
        id: e("review"), from: "COO", to: "CEO", subject: "Helios review",
        body: "Two weeks after the Helios kickoff, let's hold a review on {@helios_kickoff+2w @14:00-15:00}.",
        depends_on: [{ email: e("kickoff"), type: "date" }],
        answer: { ops: [{ create: "Helios review", kind: "event", on: { eq: "@helios_kickoff+2w @14:00-15:00" } }] },
      },
      {
        // 3. A RESCHEDULE. Conflicts are authored as a later email that MOVES an earlier event by name (never
        // "find an open slot"). `move` targets the obligation "Helios review" and gives its new date; it
        // inherits the event kind from the create. The new date is one week (not two) after the kickoff.
        id: e("review-moved"), from: "BOARD", to: "CEO", subject: "Re: Helios review, moved up a week",
        body: "The board needs the Helios review moved up a week, same time. The new slot is {@helios_kickoff+1w @14:00-15:00}.",
        depends_on: [{ email: e("review"), type: "date" }],
        answer: { ops: [{ move: "Helios review", on: { eq: "@helios_kickoff+1w @14:00-15:00" } }] },
      },
      {
        // 4. A TO-DO ON A DEADLINE. Not an event: no clock, and graded with `by` ("on or before") instead of
        // `eq` ("on exactly"). The assistant is correct if it lands the task on ANY day up to and including the
        // deadline. "+10bd" = ten BUSINESS days after the kickoff. Also a needle (reuses @helios_kickoff).
        id: e("filing"), from: "COO", to: "CEO", subject: "Helios board filing due",
        body: "One task for you: the Helios board filing needs to be submitted by {@helios_kickoff+10bd}. Please add it to your to-do list.",
        depends_on: [{ email: e("kickoff"), type: "date" }],
        answer: { ops: [{ create: "Helios board filing", kind: "todo", on: { by: "@helios_kickoff+10bd" } }] },
      },
      {
        // 5. MANY ACTIONS IN ONE EMAIL (what the "scaffold an action for each date" button is for). Three
        // partner windows are written once in this body as anchors; the answer creates one event per window,
        // reusing each by name. These anchors are emitted AND reused inside the SAME email, which is fine and
        // needs no `date` edge. The body bases them off the kickoff, so the only edge is a plain `static` one
        // (cross-email ordering). Each obligation has a distinct name, so the grader keeps them apart.
        id: e("partners"), from: "BIZDEV", to: "CEO", subject: "Three partner intros to schedule",
        body: "Three partners want to meet about Helios. Acme can do {!acme_slot = @helios_kickoff+3w @09:00-10:00}, Globex prefers {!globex_slot = @helios_kickoff+3w @11:00-12:00}, and Initech offered {!initech_slot = @helios_kickoff+4w @15:00-16:00}. Please get all three on the calendar.",
        depends_on: [{ email: e("kickoff"), type: "static" }],
        answer: { ops: [
          { create: "Meet with Acme", kind: "event", on: { eq: "@acme_slot" } },
          { create: "Meet with Globex", kind: "event", on: { eq: "@globex_slot" } },
          { create: "Meet with Initech", kind: "event", on: { eq: "@initech_slot" } },
        ] },
      },
      {
        // 6. A CANCEL. `cancel` takes a name only (no date, no kind — both inherited from the create). After
        // this turn, zero "Meet with Globex" events should remain; the other two partner meetings stay. The
        // edge back to the email that created it is derived for you (it's a `static` one — no date involved).
        id: e("globex-cancel"), from: "BIZDEV", to: "CEO", subject: "Re: Globex meeting cancelled",
        body: "Update: Globex had to back out this quarter. Please take the Globex meeting back off the calendar.",
        depends_on: [{ email: e("partners"), type: "static" }],
        answer: { ops: [{ cancel: "Meet with Globex" }] },
      },
      {
        // 7. ANY OF A FEW DAYS. When more than one day is genuinely acceptable, list them with `any_of` and the
        // assistant is correct if it lands on ANY one. (This is the only deliberately multi-day answer besides
        // a `by` deadline — there is no "find a free slot".) Day-level here: no clock, so it grades by day.
        id: e("offsite"), from: "COO", to: "CEO", subject: "Helios offsite, pick a day",
        body: "Let's get a Helios offsite on the books. Either {next:MON} or {next:WED} works for the team, so put it on whichever is easier.",
        depends_on: [],
        answer: { ops: [{ create: "Helios offsite", kind: "event", on: { any_of: ["next:MON", "next:WED"] } }] },
      },
      {
        // 8. NO ACTION (a distractor). An FYI the assistant should ignore: empty ops. A correct assistant does
        // NOTHING here; anything it creates is a failure. In the builder this is the "this email needs no
        // action" checkbox. A good corpus is mostly emails like this so the action emails aren't obvious.
        id: e("mockups-fyi"), from: "DESIGN", to: "CEO", subject: "Helios mockups arriving Thursday",
        body: "Heads up: FedEx is dropping the Helios mockups this Thursday. Nothing you need to do.",
        depends_on: [],
        answer: { ops: [] },
      },
    ],
  };
}
