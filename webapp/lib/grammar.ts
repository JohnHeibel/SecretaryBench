// Grammar constants + small pure helpers shared across the UI. The authoritative
// parser/evaluator is the Python sb.resolver (via /api/resolve, /api/lint) — these
// are only for building UI affordances (dropdowns, anchor lists, token scans).
import type { Answer, CorpusNode, ObjKind, Op, Verb } from "./types";

export const WEEKDAYS = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"] as const;
export const UNITS: { value: string; label: string }[] = [
  { value: "d", label: "calendar days" },
  { value: "bd", label: "business days" },
  { value: "w", label: "weeks" },
  { value: "m", label: "months" },
  { value: "y", label: "years" },
];
export const NTH = ["1", "2", "3", "4", "5", "last"] as const;

// The answer-key verb model (sb/schema.py: create/move/cancel on a named obligation).
// The dropdown folds verb + kind into one friendly choice; `kind` is only meaningful on
// create — move/cancel inherit it from the obligation's create op.
export const OP_CHOICES: { id: string; label: string; verb: Verb; kind?: ObjKind }[] = [
  { id: "create_event", label: "Create an event", verb: "create", kind: "event" },
  { id: "create_todo", label: "Create a to-do", verb: "create", kind: "todo" },
  { id: "move", label: "Move / reschedule", verb: "move" },
  { id: "cancel", label: "Cancel", verb: "cancel" },
];

export function opVerb(op: Op): Verb {
  return op.create !== undefined ? "create" : op.move !== undefined ? "move" : "cancel";
}
export function opName(op: Op): string {
  return (op[opVerb(op)] as string) ?? "";
}
export function opChoiceId(op: Op): string {
  const v = opVerb(op);
  if (v === "create") return op.kind === "todo" ? "create_todo" : "create_event";
  return v;
}

// Coerce any stored answer (a partial/missing field, or a legacy expect/forbid row) into
// the verb-model shape so the editor can never crash on bad input. A legacy entry can't be
// faithfully translated, so it collapses to no-action (ops: []) and the author re-authors —
// strictly better than a crash, and the lint gate flags the now-empty needle.
export function normalizeAnswer(a: unknown): Answer {
  const obj = (a ?? {}) as Record<string, unknown>;
  const ops = Array.isArray(obj.ops) ? (obj.ops as Op[]) : [];
  const emits = obj.emits && typeof obj.emits === "object" ? (obj.emits as Record<string, string>) : undefined;
  return emits ? { ops, emits } : { ops };
}

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
