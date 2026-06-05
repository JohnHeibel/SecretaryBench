// Ready-made scenarios authors can load as a starting point. The flagship is "Project Helios",
// the worked example in HOW_IT_WORKS.md: it exercises the whole benchmark in four emails — a fixed
// timed event, an anchor reused across a long gap (the long-horizon needle), a reschedule conflict,
// and a no-action distractor — and is built to lint clean AND solve 1.0 under the reference oracle.
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
        body: "Hi — we're kicking off Project Helios. Please put the kickoff on the calendar for {!helios_kickoff = next:THU @10:00-11:00}. Thanks!",
        depends_on: [],
        answer: { ops: [{ create: "Helios kickoff", kind: "event", on: { eq: "next:THU @10:00-11:00" } }] },
        tier: "T1", type: "action",
      },
      {
        id: e("review"), from: "COO", to: "CEO", subject: "Helios review",
        body: "Two weeks after the Helios kickoff, let's hold a review on {@helios_kickoff+2w @14:00-15:00}.",
        depends_on: [{ email: e("kickoff"), type: "date" }],
        answer: { ops: [{ create: "Helios review", kind: "event", on: { eq: "@helios_kickoff+2w @14:00-15:00" } }] },
        tier: "T3", type: "action",
      },
      {
        id: e("review-moved"), from: "BOARD", to: "CEO", subject: "Re: Helios review — moved up a week",
        body: "The board needs the Helios review moved up a week, same time — the new slot is {@helios_kickoff+1w @14:00-15:00}.",
        depends_on: [{ email: e("review"), type: "date" }],
        answer: { ops: [{ move: "Helios review", on: { eq: "@helios_kickoff+1w @14:00-15:00" } }] },
        tier: "T3", type: "action",
      },
      {
        id: e("fyi"), from: "DESIGN", to: "CEO", subject: "Helios mockups arriving Thursday",
        body: "Heads up: FedEx is dropping the Helios mockups this Thursday. Nothing you need to do.",
        depends_on: [],
        answer: { ops: [] },
        type: "no_action",
      },
    ],
  };
}
