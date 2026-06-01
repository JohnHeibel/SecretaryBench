// Grammar constants + small pure helpers shared across the UI. The authoritative
// parser/evaluator is the Python sb.resolver (via /api/resolve, /api/lint) — these
// are only for building UI affordances (dropdowns, anchor lists, token scans).
import type { CorpusNode } from "./types";

export const WEEKDAYS = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"] as const;
export const UNITS: { value: string; label: string }[] = [
  { value: "d", label: "calendar days" },
  { value: "bd", label: "business days" },
  { value: "w", label: "weeks" },
  { value: "m", label: "months" },
  { value: "y", label: "years" },
];
export const NTH = ["1", "2", "3", "4", "5", "last"] as const;
export const ACTIONS = ["create_event", "create_todo", "reschedule", "reply", "delegate"] as const;

const EMIT_IN_BODY = /\{\s*!\s*([A-Za-z_][A-Za-z0-9_]*)\s*=/g;

// Anchor names a node's emails emit, via body {!name=...} or answer.emits — used to
// populate the @anchor dropdown so authors pick real anchors, never typos.
export function anchorsInCorpus(nodes: CorpusNode[]): string[] {
  const names = new Set<string>();
  for (const node of nodes) {
    for (const email of node.emails) {
      for (const m of email.body.matchAll(EMIT_IN_BODY)) names.add(m[1]);
      for (const name of Object.keys(email.answer?.emits ?? {})) names.add(name);
    }
  }
  return [...names].sort();
}

// Split a body into literal text and {token} spans so we can render chips.
export interface BodySpan {
  text: string;
  token: boolean;
}
export function splitBody(body: string): BodySpan[] {
  const spans: BodySpan[] = [];
  const re = /\{[^{}]*\}/g;
  let last = 0;
  for (const m of body.matchAll(re)) {
    if (m.index! > last) spans.push({ text: body.slice(last, m.index), token: false });
    spans.push({ text: m[0], token: true });
    last = m.index! + m[0].length;
  }
  if (last < body.length) spans.push({ text: body.slice(last), token: false });
  return spans;
}
