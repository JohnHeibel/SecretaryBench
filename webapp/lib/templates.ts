// The one ready-made scenario authors can load as a starting point: "Project Atlas". It is a single
// storyline (one cast, one DAG) deliberately built to exercise EVERYTHING the authoring tool offers and
// to show the benchmark AT SCALE, so a new author can load one finished thread and see how every piece
// is built instead of guessing. It lints clean AND solves 1.0 under the reference oracle, standalone and
// when buried under hundreds of filler emails (verified — see scripts/checkTemplate.mts + the at-scale demo).
//
// Project Atlas is a realistic business saga: a software company shipping its new flagship product, "Atlas".
// Its payoffs land far from their setups on purpose — the two load-bearing dates are published once at the
// very top (the code freeze {!atlas_freeze} and the public launch {!atlas_launch}, ten weeks out) and emails
// near the BOTTOM reuse them. When the runner interleaves filler over a long horizon those setup->payoff gaps
// (the "spans") grow large, so the assistant must actually search the inbox to recover a date set months
// earlier — the long-horizon retrieval the benchmark exists to measure.
//
// How the inbox model works NOW (mirror it when you author): every email is written TO the CEO — the boss the
// assistant works for — so there is no recipient or Cc control; you only pick who each email is FROM. The cast
// is drawn from the standard roster (clean role names), and From/To are inbox presentation only, never graded.
//
// What it shows, end to end (every construct + every authoring-tool feature):
//    1. freeze        EVENT at an exact time; publishes @atlas_freeze + @atlas_launch                        [setup, multi-anchor]
//    2. ceo-note      CEO-SENT (From: CEO) instruction to the assistant; a needle off @atlas_launch          [CEO-sent + needle]
//    3. legal-fyi     NO action (a distractor that looks like a task)                                         [restraint]
//    4. reviews       MANY actions in one email: three readiness reviews off the freeze date                 [multi-action scaffold]
//    5. beta          a TO-DO due BY a deadline (business days after freeze) = a needle                      [todo + by + needle]
//    6. press-hold    NO action (an FYI)                                                                     [restraint]
//    7. board-demo    EVENT, a needle: the board demo two weeks after freeze                                 [needle, medium span]
//    8. demo-moved    RESCHEDULE: move the board demo up a few days (still after freeze)                     [move / reschedule]
//    9. analyst-fyi   NO action (a distractor)                                                               [restraint]
//   10. compliance    a TO-DO due BY a deadline keyed off the LAUNCH = a long-span needle; match kw          [todo + by + match + long needle]
//   11. town-hall     EVENT, ANY OF a couple of acceptable days                                              [any_of]
//   12. dry-run       a needle: the launch dry-run the week before the launch                                [needle, long span]
//   13. swag-fyi      NO action (a distractor)                                                               [restraint]
//   14. launch-day    EVENT on the launch itself = the longest-span needle (reuses @atlas_launch)            [needle, longest span]
//   15. launch-dinner EVENT, a needle: a launch-night dinner (reuses @atlas_launch) — the cancel's target    [needle + sets up cancel]
//   16. dinner-cancel CANCEL: that launch-night dinner is called off                                         [cancel chain]
//   17. retro         a needle: the launch retro two weeks after launch                                      [needle, beyond launch]
//   18. allhands-fyi  NO action (the final distractor)                                                       [restraint]
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
    cast: { CEO: "you (the boss)", VP_PRODUCT: "VP Product", VP_ENG: "VP Engineering", GC: "General Counsel", COMMS: "Communications / PR", BOARD_CHAIR: "Board Chair", IR: "Investor Relations", HR: "Head of HR" },
    emails: [
      {
        // 1. THE SETUP (multi-anchor). The two dates the whole saga hangs on are published HERE, once, as body
        // anchors: the code freeze this coming Monday, and the public launch ten weeks after it. Both are reused
        // later by emails far down the thread, so these are the deepest needles once filler is added. Like every
        // email it's from a person (VP Product) to the CEO. The freeze itself is a real event with a clock; its
        // answer reuses @atlas_freeze (same-email reuse — fine, no date edge to itself).
        id: e("freeze"), from: "VP_PRODUCT", to: "CEO", subject: "Project Atlas: code freeze Monday, launch in 10 weeks",
        body: "We're locking the Atlas launch plan. Code freeze is set for {!atlas_freeze = next:MON @09:00-10:00} — please hold that hour for the go/no-go. Public launch is targeted for {!atlas_launch = next:MON+10w}; every milestone below keys off those two dates.",
        depends_on: [],
        answer: { ops: [{ create: "Atlas code freeze", kind: "event", on: { eq: "@atlas_freeze" } }] },
      },
      {
        // 2. A CEO-SENT EMAIL. From can be the boss firing an instruction at their own assistant (From: CEO) —
        // a normal, encouraged shape, not just inbound mail. It's also a needle: it reuses @atlas_launch (+1w), so
        // it carries a `date` edge back to email 1 and lands far in the future (a generous, always-feasible window).
        id: e("ceo-note"), from: "CEO", to: "CEO", subject: "Set up my launch-week press interview",
        body: "Note for you to action: once Atlas launches, get my first press interview on the books the week after — {@atlas_launch+1w @11:00-12:00}.",
        depends_on: [{ email: e("freeze"), type: "date" }],
        answer: { ops: [{ create: "Atlas press interview", kind: "event", on: { eq: "@atlas_launch+1w @11:00-12:00" } }] },
      },
      {
        // 3. NO ACTION (looks like a task, isn't). Legal sends a heads-up that reads like a request but asks
        // for nothing schedulable. A correct assistant does nothing; the "needs no action" box = empty ops.
        id: e("legal-fyi"), from: "GC", to: "CEO", subject: "Re: Project Atlas: trademark cleared",
        body: "FYI only: the Atlas trademark cleared this morning, so marketing is unblocked on the name. No action needed from you, I just wanted you in the loop.",
        depends_on: [{ email: e("freeze"), type: "static" }],
        answer: { ops: [] },
      },
      {
        // 4. MANY ACTIONS IN ONE EMAIL (the scaffold path). Three readiness reviews, each written once as a
        // body anchor off the freeze date, then one create op apiece. The anchors are emitted AND reused in
        // THIS SAME email, so the only edge is a plain `static` one back to freeze — a `date` edge here would
        // be a self-edge and deadlock the scheduler. Distinct names keep the three apart in the grader.
        id: e("reviews"), from: "VP_ENG", to: "CEO", subject: "Re: Atlas readiness: three review sessions",
        body: "To ship clean we need three Atlas readiness reviews on the calendar. Performance review can go {!rev_perf = @atlas_freeze+3d @09:00-10:30}, the security review wants {!rev_security = @atlas_freeze+1w @13:00-15:00}, and the content/docs review is set for {!rev_content = @atlas_freeze+10d @11:00-12:00}. Please book all three.",
        depends_on: [{ email: e("freeze"), type: "static" }],
        answer: { ops: [
          { create: "Atlas performance review", kind: "event", on: { eq: "@rev_perf" } },
          { create: "Atlas security review", kind: "event", on: { eq: "@rev_security" } },
          { create: "Atlas content review", kind: "event", on: { eq: "@rev_content" } },
        ] },
      },
      {
        // 5. A TO-DO ON A DEADLINE (a needle). Not an event (no clock); graded `by` (on or before). Ten
        // business days after the freeze the beta feedback must be compiled. Reuses @atlas_freeze, so it
        // carries a `date` edge back to email 1 and is a genuine needle.
        id: e("beta"), from: "VP_PRODUCT", to: "CEO", subject: "Atlas beta feedback: compile the results",
        body: "One to-do for you: the Atlas beta feedback needs to be compiled and summarized by {@atlas_freeze+10bd}. Add it to your list so it doesn't slip.",
        depends_on: [{ email: e("freeze"), type: "date" }],
        answer: { ops: [{ create: "Atlas beta feedback", kind: "todo", on: { by: "@atlas_freeze+10bd" } }] },
      },
      {
        // 6. NO ACTION (an FYI). A comms heads-up with nothing to schedule. Pure restraint test.
        id: e("press-hold"), from: "COMMS", to: "CEO", subject: "Atlas: holding the press release",
        body: "Heads up: we're holding the Atlas press release until launch day, per legal. Nothing for you to do right now, just flagging the embargo.",
        depends_on: [{ email: e("freeze"), type: "static" }],
        answer: { ops: [] },
      },
      {
        // 7. A NEEDLE (medium span). The board demo two weeks after the freeze. An event with a clock that
        // reuses @atlas_freeze in its answer -> date edge back to email 1.
        id: e("board-demo"), from: "BOARD_CHAIR", to: "CEO", subject: "Atlas: live board demo",
        body: "The board wants a live Atlas demo two weeks after freeze. Please put it on the calendar for {@atlas_freeze+2w @15:00-16:00}.",
        depends_on: [{ email: e("freeze"), type: "date" }],
        answer: { ops: [{ create: "Atlas board demo", kind: "event", on: { eq: "@atlas_freeze+2w @15:00-16:00" } }] },
      },
      {
        // 8. A RESCHEDULE. A later email MOVES the board demo up a few days (a scheduling conflict), same time.
        // `move` targets the obligation "Atlas board demo" by name and inherits its event kind. The new date
        // still reuses @atlas_freeze, so the edge to the creator is a `date` edge. Moving EARLIER keeps the new
        // date safely after freeze, so the serve window stays feasible.
        id: e("demo-moved"), from: "BOARD_CHAIR", to: "CEO", subject: "Re: Atlas board demo, pulling it in a few days",
        body: "Scheduling's tight that week, so we're pulling the Atlas board demo in to {@atlas_freeze+11d @15:00-16:00}, same hour.",
        depends_on: [{ email: e("board-demo"), type: "date" }],
        answer: { ops: [{ move: "Atlas board demo", on: { eq: "@atlas_freeze+11d @15:00-16:00" } }] },
      },
      {
        // 9. NO ACTION (a distractor). An analyst-interest FYI; nothing to book.
        id: e("analyst-fyi"), from: "IR", to: "CEO", subject: "Atlas: analyst questions trickling in",
        body: "A few analysts are asking about Atlas already. We're routing everything through IR and have it handled. No action needed.",
        depends_on: [{ email: e("freeze"), type: "static" }],
        answer: { ops: [] },
      },
      {
        // 10. A LONG-SPAN TO-DO NEEDLE keyed off the LAUNCH, not the freeze. The accessibility audit is due
        // five business days BEFORE the launch. Because @atlas_launch is ten weeks out, this needle's setup
        // (email 1) and payoff sit far apart -> a large span under filler. `match` overrides the keyword the
        // grader looks for to "accessibility", since that's what the calendar entry would naturally be called.
        id: e("compliance"), from: "GC", to: "CEO", subject: "Atlas: accessibility audit before launch",
        body: "Compliance: the Atlas accessibility audit has to be submitted by {@atlas_launch-5bd}. Please add it as a to-do — missing it slips the whole launch.",
        depends_on: [{ email: e("freeze"), type: "date" }],
        answer: { ops: [{ create: "Atlas accessibility audit", kind: "todo", match: ["accessibility"], on: { by: "@atlas_launch-5bd" } }] },
      },
      {
        // 11. ANY OF a few days. An all-staff launch briefing; either of two mornings works, so the assistant
        // is correct landing on EITHER. Day-level any_of (the dates carry no clock here).
        id: e("town-hall"), from: "HR", to: "CEO", subject: "Atlas: all-staff launch briefing, pick a morning",
        body: "Let's get an Atlas all-staff briefing booked. Either {next:TUE+1w} or {next:THU+1w} works for the team, so put it on whichever fits your morning better.",
        depends_on: [{ email: e("freeze"), type: "static" }],
        answer: { ops: [{ create: "Atlas all-staff briefing", kind: "event", on: { any_of: ["next:TUE+1w", "next:THU+1w"] } }] },
      },
      {
        // 12. A LONG-SPAN NEEDLE. The launch dry-run the week before the launch. Reuses @atlas_launch (-1w)
        // -> a date edge back to the very first email, and a big span once filler spaces them apart.
        id: e("dry-run"), from: "VP_ENG", to: "CEO", subject: "Atlas: launch dry-run, week of launch",
        body: "Let's lock the Atlas launch dry-run for the week before launch: {@atlas_launch-1w @10:00-11:30}. That gives the team a clean week to fix anything that breaks.",
        depends_on: [{ email: e("freeze"), type: "date" }],
        answer: { ops: [{ create: "Atlas launch dry-run", kind: "event", on: { eq: "@atlas_launch-1w @10:00-11:30" } }] },
      },
      {
        // 13. NO ACTION (a distractor). A swag/logistics FYI; nothing to schedule.
        id: e("swag-fyi"), from: "HR", to: "CEO", subject: "Atlas: launch swag for the team",
        body: "We've ordered Atlas launch swag for the whole team to land before launch day. All handled on our side, nothing for you to action.",
        depends_on: [{ email: e("freeze"), type: "static" }],
        answer: { ops: [] },
      },
      {
        // 14. THE LONGEST-SPAN NEEDLE. The launch itself, reusing @atlas_launch directly (the date published in
        // email 1, ten weeks earlier). Under filler this is the deepest setup->payoff gap in the corpus — the
        // headline long-horizon test. An event with a clock.
        id: e("launch-day"), from: "VP_PRODUCT", to: "CEO", subject: "Atlas: launch-day keynote",
        body: "Launch day is here: please block the Atlas launch keynote for {@atlas_launch @09:00-10:00}. This is the one that matters.",
        depends_on: [{ email: e("freeze"), type: "date" }],
        answer: { ops: [{ create: "Atlas launch keynote", kind: "event", on: { eq: "@atlas_launch @09:00-10:00" } }] },
      },
      {
        // 15. The create the cancel (email 16) targets — a separate email so the cancel has a real prior
        // obligation to remove. Reuses @atlas_launch (an anchor from ANOTHER email), so it's a true cross-email
        // needle and carries a `date` edge to email 1.
        id: e("launch-dinner"), from: "COMMS", to: "CEO", subject: "Atlas: launch-night team dinner",
        body: "Let's celebrate the Atlas launch with a team dinner the evening of launch day: {@atlas_launch @19:00-21:00}. Penciling it in.",
        depends_on: [{ email: e("freeze"), type: "date" }],
        answer: { ops: [{ create: "Atlas launch dinner", kind: "event", on: { eq: "@atlas_launch @19:00-21:00" } }] },
      },
      {
        // 16. A CANCEL. The launch-night dinner is called off. `cancel` names the obligation only (no date, no
        // kind — both inherited from its create above). After this, zero "Atlas launch dinner" events remain.
        id: e("dinner-cancel"), from: "COMMS", to: "CEO", subject: "Re: Atlas launch dinner, let's not",
        body: "Change of plan: with the team spread across offices we're going to skip the Atlas launch dinner this time. If it's on your calendar, please take it off.",
        depends_on: [{ email: e("launch-dinner"), type: "static" }],
        answer: { ops: [{ cancel: "Atlas launch dinner" }] },
      },
      {
        // 17. A NEEDLE BEYOND THE LAUNCH. The launch retro two weeks after launch. Reuses @atlas_launch+2w -> a
        // date edge to email 1, and one of the latest payoffs in the saga.
        id: e("retro"), from: "VP_PRODUCT", to: "CEO", subject: "Atlas: launch retro",
        body: "Once the dust settles, let's do an Atlas launch retro two weeks after launch: {@atlas_launch+2w @14:00-15:00}.",
        depends_on: [{ email: e("freeze"), type: "date" }],
        answer: { ops: [{ create: "Atlas launch retro", kind: "event", on: { eq: "@atlas_launch+2w @14:00-15:00" } }] },
      },
      {
        // 18. NO ACTION (the final distractor). A celebratory all-hands recap with nothing to do.
        id: e("allhands-fyi"), from: "HR", to: "CEO", subject: "Atlas: recap going out to all-hands",
        body: "We'll fold the Atlas launch milestones into the next all-hands deck. Comms owns it end to end, so there's nothing you need to do here.",
        depends_on: [{ email: e("freeze"), type: "static" }],
        answer: { ops: [] },
      },
    ],
  };
}
