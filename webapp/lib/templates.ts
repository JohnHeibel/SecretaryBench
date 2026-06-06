// The one ready-made scenario authors can load as a starting point: "Project Atlas". It is a single
// storyline with a standard roster, sender-only emails, and a compact product-launch arc that is small
// enough to narrate in a how-to video. It lints clean and solves 1.0 under the reference oracle.
//
// Project Atlas is a software company shipping its new flagship product, Atlas. The first email publishes
// the two load-bearing dates, code freeze and public launch, and later emails reuse those anchors. That
// creates the long-horizon retrieval test: the assistant has to recover dates that were set earlier in the
// inbox instead of guessing from the current email alone.
//
// How the inbox model works now: every email is written to the CEO, so authors only choose who each email is
// from. The cast uses the standard roster keys. From, To, and Cc are inbox presentation only, never graded.
//
// What it shows, end to end:
//    1. freeze        EVENT at an exact time; publishes @atlas_freeze and @atlas_launch       [setup, multi-anchor]
//    2. ceo-note      CEO-sent instruction; a needle off @atlas_launch                       [CEO-sent, needle]
//    3. legal-fyi     NO action, a legal update that looks tempting                          [restraint]
//    4. beta          TO-DO due BY a timed deadline after freeze                             [todo, timed by, needle]
//    5. board-demo    EVENT two weeks after freeze                                           [needle, medium span]
//    6. demo-moved    RESCHEDULE the board demo                                              [move]
//    7. launch-dinner EVENT on launch night, reusing @atlas_launch                           [long needle, cancel target]
//    8. dinner-cancel CANCEL that launch dinner                                              [cancel chain]
//
// Two rules worth internalizing from this example:
//   - The body and the answer point at the same date. A date written once as {!name = expr} in a body is
//     reused in the answer as @name, so the rendered email and answer key cannot drift.
//   - "Comes after" vs "reuses an earlier date" edges: a reschedule or needle that reuses an earlier date
//     carries a `date` edge. Plain ordering uses `static`. The app derives these edges for authors, but this
//     raw template writes them out so the JSON lints on its own.
import type { CorpusNode } from "./types";

// `nodeId` lets the caller pick a non-colliding storyline id; email ids are prefixed with it so the
// dependency edges and exported corpus stay readable.
export function projectAtlasNode(nodeId = "project_atlas"): CorpusNode {
  const e = (slug: string) => `${nodeId}.${slug}`;
  return {
    id: nodeId,
    cast: { CEO: "you (the boss)", VP_PRODUCT: "VP Product", VP_ENG: "VP Engineering", GC: "General Counsel", COMMS: "Communications / PR", BOARD_CHAIR: "Board Chair" },
    emails: [
      {
        // 1. Setup: publish the two dates that the rest of the story reuses. The freeze itself is an event
        // with a clock, and the launch anchor stays day-level so later emails can attach their own times.
        id: e("freeze"), from: "VP_PRODUCT", to: "CEO", subject: "Project Atlas: code freeze Monday, launch in 10 weeks",
        body: "We're locking the Atlas launch plan. Code freeze is set for {!atlas_freeze = next:MON @09:00-10:00}; please hold that hour for the go/no-go. Public launch is targeted for {!atlas_launch = next:MON+10w}.",
        depends_on: [],
        answer: { ops: [{ create: "Atlas code freeze", kind: "event", on: { eq: "@atlas_freeze" } }] },
      },
      {
        // 2. CEO-sent email: the boss can send their own assistant an instruction. It is also a long needle
        // because the date is built from the launch anchor published in the first email.
        id: e("ceo-note"), from: "CEO", to: "CEO", subject: "Set up my launch-week press interview",
        body: "Note for you to action: once Atlas launches, get my first press interview on the books the week after at {@atlas_launch+1w @11:00-12:00}.",
        depends_on: [{ email: e("freeze"), type: "date" }],
        answer: { ops: [{ create: "Atlas press interview", kind: "event", on: { eq: "@atlas_launch+1w @11:00-12:00" } }] },
      },
      {
        // 3. No action: a useful legal update that should not create anything.
        id: e("legal-fyi"), from: "GC", to: "CEO", subject: "Re: Project Atlas trademark cleared",
        body: "FYI only: the Atlas trademark cleared this morning, so marketing is unblocked on the name. No action needed from you, I just wanted you in the loop.",
        depends_on: [{ email: e("freeze"), type: "static" }],
        answer: { ops: [] },
      },
      {
        // 4. Timed deadline: a to-do due by a clock cutoff. The grader compares the created object's start
        // to the 5 PM cutoff, while the scheduler still uses the date as the serve-by window.
        id: e("beta"), from: "VP_PRODUCT", to: "CEO", subject: "Atlas beta feedback: compile the results",
        body: "One to-do for you: the Atlas beta feedback needs to be compiled and summarized by {@atlas_freeze+10bd @17:00}. Add it to your list so it doesn't slip.",
        depends_on: [{ email: e("freeze"), type: "date" }],
        answer: { ops: [{ create: "Atlas beta feedback", kind: "todo", on: { by: "@atlas_freeze+10bd @17:00" } }] },
      },
      {
        // 5. Board demo: a medium-span needle that creates an event off the freeze date.
        id: e("board-demo"), from: "BOARD_CHAIR", to: "CEO", subject: "Atlas live board demo",
        body: "The board wants a live Atlas demo two weeks after freeze. Please put it on the calendar for {@atlas_freeze+2w @15:00-16:00}.",
        depends_on: [{ email: e("freeze"), type: "date" }],
        answer: { ops: [{ create: "Atlas board demo", kind: "event", on: { eq: "@atlas_freeze+2w @15:00-16:00" } }] },
      },
      {
        // 6. Move: a later email reschedules the board demo by naming the existing obligation.
        id: e("demo-moved"), from: "BOARD_CHAIR", to: "CEO", subject: "Re: Atlas board demo, pulling it in",
        body: "Scheduling is tight that week, so we are pulling the Atlas board demo in to {@atlas_freeze+11d @15:00-16:00}, same hour.",
        depends_on: [{ email: e("board-demo"), type: "date" }],
        answer: { ops: [{ move: "Atlas board demo", on: { eq: "@atlas_freeze+11d @15:00-16:00" } }] },
      },
      {
        // 7. Launch dinner: the longest-span needle, reusing the launch anchor ten weeks after setup. This
        // creates the object that the final email cancels.
        id: e("launch-dinner"), from: "COMMS", to: "CEO", subject: "Atlas launch-night team dinner",
        body: "Let's celebrate the Atlas launch with a team dinner the evening of launch day: {@atlas_launch @19:00-21:00}. Penciling it in.",
        depends_on: [{ email: e("freeze"), type: "date" }],
        answer: { ops: [{ create: "Atlas launch dinner", kind: "event", on: { eq: "@atlas_launch @19:00-21:00" } }] },
      },
      {
        // 8. Cancel: remove the dinner created above. The cancel names the obligation only.
        id: e("dinner-cancel"), from: "COMMS", to: "CEO", subject: "Re: Atlas launch dinner, let's not",
        body: "Change of plan: with the team spread across offices we are going to skip the Atlas launch dinner this time. If it's on your calendar, please take it off.",
        depends_on: [{ email: e("launch-dinner"), type: "static" }],
        answer: { ops: [{ cancel: "Atlas launch dinner" }] },
      },
    ],
  };
}
