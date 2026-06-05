// The one ready-made scenario authors can load as a starting point: "Project Atlas". It is a single
// storyline (one cast, one DAG) deliberately built to exercise EVERYTHING the authoring tool offers and
// to show the benchmark AT SCALE, so a new author can load one finished thread and see how every piece
// is built instead of guessing. It lints clean AND solves 1.0 under the reference oracle, standalone and
// when buried under hundreds of filler emails (verified — see scripts/checkTemplate.mts + the at-scale demo).
//
// Project Atlas is a realistic business saga: a SaaS company ("yours") acquiring a competitor, Northwind.
// Its payoffs land far from their setups on purpose — the two load-bearing dates are published once at the
// very top (the signing {!atlas_signing} and the close {!atlas_close}, twelve weeks out) and emails near
// the BOTTOM reuse them. When the runner interleaves filler over a long horizon those setup->payoff gaps
// (the "spans") grow large, so the assistant must actually search the inbox to recover a date set months
// earlier — the long-horizon retrieval the benchmark exists to measure.
//
// What it shows, end to end (every construct + every authoring-tool feature):
//    1. signing       EVENT at an exact time; publishes @atlas_signing + @atlas_close; multi-recipient To + Cc  [setup + To/Cc]
//    2. ceo-1on1      CEO-SENT (From: CEO) instruction to the assistant; a needle off @atlas_close             [CEO-sent + needle]
//    3. legal-fyi     NO action (a distractor that looks like a task)                                           [restraint]
//    4. diligence     MANY actions in one email: three diligence sessions off the signing date                 [multi-action scaffold]
//    5. data-room     a TO-DO due BY a deadline (business days after signing) = a needle                       [todo + by + needle]
//    6. press-hold    NO action (an FYI)                                                                       [restraint]
//    7. board-vote    EVENT, a needle: the board vote two weeks after signing                                  [needle, medium span]
//    8. vote-moved    RESCHEDULE: move the board vote up a few days (still after signing)                      [move / reschedule]
//    9. analyst-fyi   NO action (a distractor)                                                                 [restraint]
//   10. reg-filing    a TO-DO due BY a regulatory deadline keyed off the CLOSE = a long-span needle; match kw  [todo + by + match + long needle]
//   11. town-hall     EVENT, ANY OF a couple of acceptable days                                                [any_of]
//   12. integration   a needle: the integration sync the week before the close                                [needle, long span]
//   13. swag-fyi      NO action (a distractor)                                                                 [restraint]
//   14. close-day     EVENT on the close itself = the longest-span needle (reuses @atlas_close)                [needle, longest span]
//   15. close-dinner  EVENT, a needle: a close-night dinner (reuses @atlas_close) — the cancel's target       [needle + sets up cancel]
//   16. dinner-cancel CANCEL: that close-night dinner is called off                                            [cancel chain]
//   17. retro         a needle: the retro two weeks after the close                                            [needle, beyond close]
//   18. allhands-fyi  NO action (the final distractor)                                                         [restraint]
//
// Two rules worth internalizing from this example:
//   - The body and the answer point at the SAME date. A date written once as {!name = expr} in a body is
//     reused in the answer as @name, so they can never disagree. (That's why the answer key isn't "a
//     different tool" from the email — it's the same dates, named.)
//   - "comes after" vs "reuses an earlier date (needle)" edges: a reschedule/needle that reuses an earlier
//     DATE carries a `date` edge (the grader derives a serve-by window from it). Plain ordering is `static`.
//     The app wires the `date` edge for you whenever an answer reuses an @anchor; you rarely set it by hand.
//
// Self-emit note (the gotcha): emails 1 and 4 both write anchors in their body AND reuse them in their OWN
// answer. Same-email reuse needs NO date edge (a date edge to yourself deadlocks the scheduler), so those
// keep only static cross-email edges. Every CROSS-email needle carries a `date` edge back to the email that
// published the anchor — the app derives these for you, written out explicitly here so the raw JSON lints
// on its own.
import type { CorpusNode } from "./types";

// `nodeId` lets the caller pick a non-colliding storyline id; email ids are prefixed with it so the
// dependency edges (and the eventual corpus/nodes/<id>.json) stay readable.
export function projectAtlasNode(nodeId = "project_atlas"): CorpusNode {
  const e = (slug: string) => `${nodeId}.${slug}`;
  return {
    id: nodeId,
    cast: { CEO: "you", CORPDEV: "Sam Okafor (Corp Dev)", LEGAL: "Tara Whitfield (General Counsel)", FINANCE: "Ng Wei Jie (CFO)", BOARD: "Board office", COMMS: "Priya Anand (Comms)", PEOPLE: "Devon Marsh (People Ops)" },
    emails: [
      {
        // 1. THE SETUP (multi-anchor + To/Cc). The two dates the whole saga hangs on are published HERE, once,
        // as body anchors: the signing this coming Friday, and the planned close twelve weeks after it. Both are
        // reused later by emails far down the thread, so these are the deepest needles once filler is added. It
        // is addressed to several people (multi-recipient To) and copies legal+finance (Cc) — To/Cc are inbox
        // presentation only and never graded, so they don't change the score. The signing itself is a real event
        // with a clock; its answer reuses @atlas_signing (same-email reuse — fine, no date edge to itself).
        id: e("signing"), from: "CORPDEV", to: ["CEO", "BOARD"], cc: ["LEGAL", "FINANCE"], subject: "Project Atlas: signing Friday, target close in 12 weeks",
        body: "We're a go on Project Atlas (the Northwind acquisition). Signing is locked for {!atlas_signing = next:FRI @10:00-11:00} — please block that hour. We are targeting close on {!atlas_close = next:FRI+12w}; everything downstream keys off those two dates.",
        depends_on: [],
        answer: { ops: [{ create: "Atlas signing", kind: "event", on: { eq: "@atlas_signing" } }] },
      },
      {
        // 2. A CEO-SENT EMAIL. From can be the boss firing an instruction at their own assistant (From: CEO) —
        // a normal, encouraged shape, not just inbound mail. It's also a needle: it reuses @atlas_close (+1w), so
        // it carries a `date` edge back to email 1 and lands far in the future (a generous, always-feasible window).
        id: e("ceo-1on1"), from: "CEO", to: "CEO", subject: "Set up my 1:1 with Northwind's CEO after close",
        body: "Note for you to action: once Atlas closes, get my first 1:1 with Northwind's CEO on the books the week after — {@atlas_close+1w @11:00-12:00}.",
        depends_on: [{ email: e("signing"), type: "date" }],
        answer: { ops: [{ create: "Northwind 1:1", kind: "event", on: { eq: "@atlas_close+1w @11:00-12:00" } }] },
      },
      {
        // 3. NO ACTION (looks like a task, isn't). Legal sends a heads-up that reads like a request but asks
        // for nothing schedulable. A correct assistant does nothing; the "needs no action" box = empty ops.
        id: e("legal-fyi"), from: "LEGAL", to: "CEO", subject: "Re: Project Atlas: NDA countersigned",
        body: "FYI only: Northwind countersigned the NDA this morning, so we're clear to share the diligence materials. No action needed from you, I just wanted you in the loop.",
        depends_on: [{ email: e("signing"), type: "static" }],
        answer: { ops: [] },
      },
      {
        // 4. MANY ACTIONS IN ONE EMAIL (the scaffold path). Three diligence sessions, each written once as a
        // body anchor off the signing date, then one create op apiece. The anchors are emitted AND reused in
        // THIS SAME email, so the only edge is a plain `static` one back to signing — a `date` edge here would
        // be a self-edge and deadlock the scheduler. Distinct names keep the three apart in the grader.
        id: e("diligence"), from: "FINANCE", to: "CEO", subject: "Re: Atlas diligence: three working sessions",
        body: "To hit the close we need three Atlas diligence sessions on the calendar. Finance can do {!dd_finance = @atlas_signing+3d @09:00-10:30}, the tech review wants {!dd_tech = @atlas_signing+1w @13:00-15:00}, and HR/people due-diligence is set for {!dd_people = @atlas_signing+10d @11:00-12:00}. Please book all three.",
        depends_on: [{ email: e("signing"), type: "static" }],
        answer: { ops: [
          { create: "Atlas finance diligence", kind: "event", on: { eq: "@dd_finance" } },
          { create: "Atlas tech diligence", kind: "event", on: { eq: "@dd_tech" } },
          { create: "Atlas people diligence", kind: "event", on: { eq: "@dd_people" } },
        ] },
      },
      {
        // 5. A TO-DO ON A DEADLINE (a needle). Not an event (no clock); graded `by` (on or before). Ten
        // business days after the signing the data room must be fully populated. Reuses @atlas_signing, so it
        // carries a `date` edge back to email 1 and is a genuine needle.
        id: e("data-room"), from: "CORPDEV", to: "CEO", subject: "Atlas data room: finish populating",
        body: "One to-do for you: the Atlas data room needs to be fully populated by {@atlas_signing+10bd}. Add it to your list so it doesn't slip.",
        depends_on: [{ email: e("signing"), type: "date" }],
        answer: { ops: [{ create: "Atlas data room", kind: "todo", on: { by: "@atlas_signing+10bd" } }] },
      },
      {
        // 6. NO ACTION (an FYI). A comms heads-up with nothing to schedule. Pure restraint test.
        id: e("press-hold"), from: "COMMS", to: "CEO", subject: "Atlas: holding the press release",
        body: "Heads up: we're holding the Atlas press release until after close, per legal. Nothing for you to do right now, just flagging the embargo.",
        depends_on: [{ email: e("signing"), type: "static" }],
        answer: { ops: [] },
      },
      {
        // 7. A NEEDLE (medium span). The board vote two weeks after the signing. An event with a clock that
        // reuses @atlas_signing in its answer -> date edge back to email 1.
        id: e("board-vote"), from: "BOARD", to: "CEO", subject: "Atlas: board ratification vote",
        body: "The board will ratify the Atlas deal two weeks after signing. Please put the vote on the calendar for {@atlas_signing+2w @15:00-16:00}.",
        depends_on: [{ email: e("signing"), type: "date" }],
        answer: { ops: [{ create: "Atlas board vote", kind: "event", on: { eq: "@atlas_signing+2w @15:00-16:00" } }] },
      },
      {
        // 8. A RESCHEDULE. A later email MOVES the board vote up a few days (a quorum issue), same time. `move`
        // targets the obligation "Atlas board vote" by name and inherits its event kind. The new date still
        // reuses @atlas_signing, so the edge to the creator is a `date` edge. Moving EARLIER keeps the new date
        // safely after signing, so the serve window stays feasible.
        id: e("vote-moved"), from: "BOARD", to: "CEO", subject: "Re: Atlas board vote, pulling it in a few days",
        body: "Quorum's tight that Friday, so we're pulling the Atlas board vote in to {@atlas_signing+11d @15:00-16:00}, same hour.",
        depends_on: [{ email: e("board-vote"), type: "date" }],
        answer: { ops: [{ move: "Atlas board vote", on: { eq: "@atlas_signing+11d @15:00-16:00" } }] },
      },
      {
        // 9. NO ACTION (a distractor). An analyst-call FYI; nothing to book.
        id: e("analyst-fyi"), from: "COMMS", to: "CEO", subject: "Atlas: analyst questions trickling in",
        body: "A few analysts are asking about Atlas already. We're routing everything through IR and have it handled. No action needed.",
        depends_on: [{ email: e("signing"), type: "static" }],
        answer: { ops: [] },
      },
      {
        // 10. A LONG-SPAN TO-DO NEEDLE keyed off the CLOSE, not the signing. The HSR / regulatory filing is due
        // five business days BEFORE the close. Because @atlas_close is twelve weeks out, this needle's setup
        // (email 1) and payoff sit far apart -> a large span under filler. `match` overrides the title to "HSR"
        // since that's what the calendar entry would naturally be called.
        id: e("reg-filing"), from: "LEGAL", to: "CEO", subject: "Atlas: HSR filing before close",
        body: "Regulatory: our HSR filing for Atlas has to be submitted by {@atlas_close-5bd}. Please add it as a to-do — missing it slips the whole close.",
        depends_on: [{ email: e("signing"), type: "date" }],
        answer: { ops: [{ create: "Atlas HSR filing", kind: "todo", match: ["HSR"], on: { by: "@atlas_close-5bd" } }] },
      },
      {
        // 11. ANY OF a few days. An employee town hall to brief the team; either of two mornings works, so the
        // assistant is correct landing on EITHER. Day-level any_of (the dates carry no clock here).
        id: e("town-hall"), from: "PEOPLE", to: "CEO", subject: "Atlas: employee town hall, pick a morning",
        body: "Let's get an Atlas all-staff town hall booked. Either {next:TUE+1w} or {next:THU+1w} works for the team, so put it on whichever fits your morning better.",
        depends_on: [{ email: e("signing"), type: "static" }],
        answer: { ops: [{ create: "Atlas town hall", kind: "event", on: { any_of: ["next:TUE+1w", "next:THU+1w"] } }] },
      },
      {
        // 12. A LONG-SPAN NEEDLE. The integration kickoff sync the week before the close. Reuses @atlas_close
        // (-1w) -> a date edge back to the very first email, and a big span once filler spaces them apart.
        id: e("integration"), from: "CORPDEV", to: "CEO", subject: "Atlas: integration sync week of close",
        body: "Let's lock the Atlas integration kickoff for the week before close: {@atlas_close-1w @10:00-11:30}. That gives the teams a clean week to stage everything.",
        depends_on: [{ email: e("signing"), type: "date" }],
        answer: { ops: [{ create: "Atlas integration sync", kind: "event", on: { eq: "@atlas_close-1w @10:00-11:30" } }] },
      },
      {
        // 13. NO ACTION (a distractor). A swag/logistics FYI; nothing to schedule.
        id: e("swag-fyi"), from: "PEOPLE", to: "CEO", subject: "Atlas: welcome kits for the Northwind team",
        body: "We've ordered welcome kits for the Northwind folks to land on day one. All handled on our side, nothing for you to action.",
        depends_on: [{ email: e("signing"), type: "static" }],
        answer: { ops: [] },
      },
      {
        // 14. THE LONGEST-SPAN NEEDLE. The close itself, reusing @atlas_close directly (the date published in
        // email 1, twelve weeks earlier). Under filler this is the deepest setup->payoff gap in the corpus —
        // the headline long-horizon test. An event with a clock.
        id: e("close-day"), from: "CORPDEV", to: "CEO", subject: "Atlas: close-day signing ceremony",
        body: "Close day is here: please block the Atlas close signing ceremony for {@atlas_close @09:00-10:00}. This is the one that matters.",
        depends_on: [{ email: e("signing"), type: "date" }],
        answer: { ops: [{ create: "Atlas close ceremony", kind: "event", on: { eq: "@atlas_close @09:00-10:00" } }] },
      },
      {
        // 15. The create the cancel (email 16) targets — a separate email so the cancel has a real prior
        // obligation to remove. Reuses @atlas_close (an anchor from ANOTHER email), so it's a true cross-email
        // needle and carries a `date` edge to email 1.
        id: e("close-dinner"), from: "COMMS", to: "CEO", subject: "Atlas: close-night team dinner",
        body: "Let's celebrate the Atlas close with a team dinner the evening of the close: {@atlas_close @19:00-21:00}. Penciling it in.",
        depends_on: [{ email: e("signing"), type: "date" }],
        answer: { ops: [{ create: "Atlas close dinner", kind: "event", on: { eq: "@atlas_close @19:00-21:00" } }] },
      },
      {
        // 16. A CANCEL. The close-night dinner is called off. `cancel` names the obligation only (no date, no
        // kind — both inherited from its create above). After this, zero "Atlas close dinner" events remain.
        id: e("dinner-cancel"), from: "COMMS", to: "CEO", subject: "Re: Atlas close dinner, let's not",
        body: "Change of plan: with the team spread across offices we're going to skip the Atlas close dinner this time. If it's on your calendar, please take it off.",
        depends_on: [{ email: e("close-dinner"), type: "static" }],
        answer: { ops: [{ cancel: "Atlas close dinner" }] },
      },
      {
        // 17. A NEEDLE BEYOND THE CLOSE. The deal retro two weeks after close. Reuses @atlas_close+2w -> a date
        // edge to email 1, and one of the latest payoffs in the saga.
        id: e("retro"), from: "CORPDEV", to: "CEO", subject: "Atlas: deal retro after close",
        body: "Once the dust settles, let's do an Atlas deal retro two weeks after close: {@atlas_close+2w @14:00-15:00}.",
        depends_on: [{ email: e("signing"), type: "date" }],
        answer: { ops: [{ create: "Atlas retro", kind: "event", on: { eq: "@atlas_close+2w @14:00-15:00" } }] },
      },
      {
        // 18. NO ACTION (the final distractor). A celebratory all-hands recap with nothing to do.
        id: e("allhands-fyi"), from: "PEOPLE", to: "CEO", subject: "Atlas: recap going out to all-hands",
        body: "We'll fold the Atlas milestones into the next all-hands deck. Comms owns it end to end, so there's nothing you need to do here.",
        depends_on: [{ email: e("signing"), type: "static" }],
        answer: { ops: [] },
      },
    ],
  };
}
