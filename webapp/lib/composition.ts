// Corpus composition against AUTHORING_GUIDE.md §2 (email-type mix) and §3 (difficulty
// tiers). The guide is fundamentally about PROPORTIONS — what share of the inbox is junk
// vs no-action vs action, and how action emails split across T1/T2/T3 — so this turns those
// target bands into something the author can see live while writing. Pure functions only;
// the CompositionPanel renders what these return.
import type { CorpusNode, Email, EmailType, Tier } from "./types";

export type EmailRole = EmailType;   // alias: an email's role IS its type

// Shared palette + labels so the panel, the sidebar badge, and the editor's role control
// all read the same. `short` is the chip/button label; `label` the full word.
export const ROLE_META: Record<EmailRole, { label: string; short: string; badge: string; bar: string; dot: string }> = {
  action:    { label: "Action",     short: "Action", badge: "bg-sky-600/30 text-sky-200",       bar: "bg-sky-500",    dot: "bg-sky-400" },
  no_action: { label: "No-action",  short: "FYI",    badge: "bg-indigo-600/30 text-indigo-200", bar: "bg-indigo-500", dot: "bg-indigo-400" },
  junk:      { label: "Junk",       short: "Junk",   badge: "bg-zinc-600/40 text-zinc-300",     bar: "bg-zinc-500",   dot: "bg-zinc-400" },
};

export const ROLE_HINT: Record<EmailRole, string> = {
  action: "Action — the assistant must schedule, add, move, or cancel something (guide target 75–85%).",
  no_action: "No-action — looks like it might need an action but doesn't; tests recognition (10–15%).",
  junk: "Junk — unrelated noise, safe to ignore entirely; tests filtering (5–10%).",
};

// Classify by the guide taxonomy. A non-empty answer MEANS the email requires action, so ops
// win over any stale `type` tag; an empty answer is a no-action test unless tagged junk. This
// keeps the panel honest even if the stored tag and the answer key ever disagree.
export function classifyEmail(e: Email): EmailRole {
  if ((e.answer?.ops?.length ?? 0) > 0) return "action";
  return e.type === "junk" ? "junk" : "no_action";
}

export interface Band { lo: number; hi: number; }
// §2 — share of ALL emails.
export const ROLE_TARGET: Record<EmailRole, Band> = {
  junk:      { lo: 5,  hi: 10 },
  no_action: { lo: 10, hi: 15 },
  action:    { lo: 75, hi: 85 },
};
// §3 — share of ACTION emails (guide ≈ 30 / 40 / 30, ±5 band).
export const TIER_TARGET: Record<Tier, Band> = {
  T1: { lo: 25, hi: 35 },
  T2: { lo: 35, hi: 45 },
  T3: { lo: 25, hi: 35 },
};

// Below this many emails the percentages are too noisy to judge against a band, so the panel
// shows them but withholds the pass/fail colouring.
export const MIN_ASSESS = 10;

export interface RoleStat { role: EmailRole; count: number; pct: number; target: Band; inBand: boolean; }
export interface TierStat { tier: Tier; count: number; pct: number; target: Band; inBand: boolean; }
export interface Composition {
  total: number;
  roles: RoleStat[];            // action, no_action, junk — always all three, in that order
  actionTotal: number;
  tiers: TierStat[];            // T1, T2, T3 — share of action emails
  untaggedAction: number;      // action emails with no difficulty tier set
}

const pctOf = (n: number, d: number) => d === 0 ? 0 : (n / d) * 100;
const within = (pct: number, b: Band) => pct >= b.lo && pct <= b.hi;

export function computeComposition(nodes: CorpusNode[]): Composition {
  const emails = nodes.flatMap((n) => n.emails);
  const total = emails.length;

  const roleCount: Record<EmailRole, number> = { action: 0, no_action: 0, junk: 0 };
  for (const e of emails) roleCount[classifyEmail(e)]++;
  const roles: RoleStat[] = (["action", "no_action", "junk"] as EmailRole[]).map((role) => {
    const pct = pctOf(roleCount[role], total);
    return { role, count: roleCount[role], pct, target: ROLE_TARGET[role], inBand: within(pct, ROLE_TARGET[role]) };
  });

  const actionEmails = emails.filter((e) => classifyEmail(e) === "action");
  const actionTotal = actionEmails.length;
  const tierCount: Record<Tier, number> = { T1: 0, T2: 0, T3: 0 };
  let untaggedAction = 0;
  for (const e of actionEmails) { if (e.tier) tierCount[e.tier]++; else untaggedAction++; }
  const tiers: TierStat[] = (["T1", "T2", "T3"] as Tier[]).map((tier) => {
    const pct = pctOf(tierCount[tier], actionTotal);
    return { tier, count: tierCount[tier], pct, target: TIER_TARGET[tier], inBand: within(pct, TIER_TARGET[tier]) };
  });

  return { total, roles, actionTotal, tiers, untaggedAction };
}
