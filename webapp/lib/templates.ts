// Ready-made scenarios authors can load as a starting point. The flagship is "Project Helios",
// the worked example in HOW_IT_WORKS.md: it exercises the whole benchmark in five emails — a fixed
// timed event, an anchor reused across a long gap (the long-horizon needle), a reschedule conflict,
// a to-do graded against a due-date deadline, and a no-action distractor — and is built to lint
// clean AND solve 1.0 under the reference oracle.
// Every date lives in a {token} so the body and the answer key resolve to the same instant.
import type { CorpusNode } from "./types";

// `nodeId` lets the caller pick a non-colliding storyline id; email ids are prefixed with it so the
// dependency edges (and the eventual corpus/nodes/<id>.json) stay readable.
export function projectHeliosNode(nodeId = "project_helios"): CorpusNode {
  const e = (slug: string) => `${nodeId}.${slug}`;
  return {
    id: nodeId,
    cast: { CEO: "you", COO: "Dana Reyes (COO)", BOARD: "Board office", DESIGN: "Priya Nair (Design)" },
    emails: [
      {
        id: e("kickoff"), from: "COO", to: "CEO", subject: "Helios kickoff",
        body: "Hi, we're kicking off Project Helios. Please put the kickoff on the calendar for {!helios_kickoff = next:THU @10:00-11:00}. Thanks!",
        depends_on: [],
        answer: { ops: [{ create: "Helios kickoff", kind: "event", on: { eq: "next:THU @10:00-11:00" } }] },
      },
      {
        id: e("review"), from: "COO", to: "CEO", subject: "Helios review",
        body: "Two weeks after the Helios kickoff, let's hold a review on {@helios_kickoff+2w @14:00-15:00}.",
        depends_on: [{ email: e("kickoff"), type: "date" }],
        answer: { ops: [{ create: "Helios review", kind: "event", on: { eq: "@helios_kickoff+2w @14:00-15:00" } }] },
      },
      {
        id: e("review-moved"), from: "BOARD", to: "CEO", subject: "Re: Helios review, moved up a week",
        body: "The board needs the Helios review moved up a week, same time. The new slot is {@helios_kickoff+1w @14:00-15:00}.",
        depends_on: [{ email: e("review"), type: "date" }],
        answer: { ops: [{ move: "Helios review", on: { eq: "@helios_kickoff+1w @14:00-15:00" } }] },
      },
      {
        // A to-do, not an event: no clock, and graded with `by` (a deadline) — the assistant gets it
        // right by landing the task on ANY day up to and including 10 business days after the kickoff.
        // Reuses the @helios_kickoff anchor, so it's also a needle (a date edge back to the kickoff).
        id: e("filing"), from: "COO", to: "CEO", subject: "Helios board filing due",
        body: "One task for you: the Helios board filing needs to be submitted by {@helios_kickoff+10bd}. Please add it to your to-do list.",
        depends_on: [{ email: e("kickoff"), type: "date" }],
        answer: { ops: [{ create: "Helios board filing", kind: "todo", on: { by: "@helios_kickoff+10bd" } }] },
      },
      {
        id: e("fyi"), from: "DESIGN", to: "CEO", subject: "Helios mockups arriving Thursday",
        body: "Heads up: FedEx is dropping the Helios mockups this Thursday. Nothing you need to do.",
        depends_on: [],
        answer: { ops: [] },
      },
    ],
  };
}
