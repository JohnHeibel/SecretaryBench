// Grammar constants + small pure helpers shared across the UI. The authoritative
// parser/evaluator is the Python sb.resolver (via /api/resolve, /api/lint) — these
// are only for building UI affordances (dropdowns, anchor lists, token scans).
import type { Answer, CorpusNode, Edge, Email, ObjKind, Op, Verb } from "./types";

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

// Anchor names THIS email publishes in its own body ({!name=...}). The answer key may reuse these
// directly (schema.py exempts a self-emitted anchor from the date-edge rule), which is what powers the
// "scaffold an action per date" button and the same-email reuse chips.
export function bodyAnchorNames(body: string): string[] {
  return [...new Set([...body.matchAll(EMIT_IN_BODY)].map((m) => m[1]))];
}

// Obligation names CREATEd anywhere in a node — what a move/cancel can target, and what a later
// email may reuse as an @anchor date. (Distinct from anchorsInCorpus, which is body/emit anchors.)
export function obligationNames(node: CorpusNode): string[] {
  const names = new Set<string>();
  for (const e of node.emails) for (const op of e.answer?.ops ?? []) if (op.create) names.add(op.create);
  return [...names].sort();
}

// name -> the email that published it (a body {!=}/emits anchor, or an obligation create-name),
// so the answer-key builder can show an author WHERE a reusable date came from ("@kickoff · from
// 'Project kickoff'"). Scans the same way emitters() does; last writer wins (lint forbids collisions).
export function anchorOrigins(nodes: CorpusNode[]): Record<string, { email: string; subject: string }> {
  const out: Record<string, { email: string; subject: string }> = {};
  for (const node of nodes) for (const e of node.emails) {
    const tag = { email: e.id, subject: e.subject };
    for (const m of e.body.matchAll(EMIT_IN_BODY)) out[m[1]] = tag;
    for (const n of Object.keys(e.answer?.emits ?? {})) out[n] = tag;
    for (const op of e.answer?.ops ?? []) if (op.create) out[op.create] = tag;
  }
  return out;
}

// name -> kind for the node's created obligations. A move/cancel op carries no `kind` in the stored
// answer (sb.schema fills it in from the create at build time), so the editor derives it from here —
// e.g. to know a reschedule targets an event and should offer a clock time.
export function obligationKinds(node: CorpusNode): Record<string, ObjKind> {
  const kinds: Record<string, ObjKind> = {};
  for (const e of node.emails) for (const op of e.answer?.ops ?? []) if (op.create && op.kind) kinds[op.create] = op.kind;
  return kinds;
}

// --- auto-wiring the needle's dependency edge ------------------------------
// A "needle" is an email whose ANSWER reuses an @anchor from an earlier email. The grader requires a
// `date` dependency edge to that earlier email (so the serve-by window is derived); sb.schema derives
// it automatically only for move/cancel. We extend that to every op (e.g. a create citing @kickoff),
// so authors never have to hand-add the edge — point at an anchor and it just works.

const ANCHOR_REF_RE = /@([A-Za-z_][A-Za-z0-9_]*)/g;

function answerAnchorRefs(answer?: Answer): Set<string> {
  const refs = new Set<string>();
  for (const op of answer?.ops ?? []) {
    if (!op.on) continue;
    for (const v of Object.values(op.on)) for (const s of Array.isArray(v) ? v : [v]) {
      if (typeof s === "string") for (const m of s.matchAll(ANCHOR_REF_RE)) refs.add(m[1]);
    }
  }
  return refs;
}

// name -> email id that emits it. Body {!name=}/answer.emits anchors are corpus-wide; obligation
// create-names are node-scoped (preferred, since a needle usually lives inside one storyline).
function emitters(nodes: CorpusNode[]): { global: Record<string, string>; byNode: Record<string, Record<string, string>> } {
  const global: Record<string, string> = {};
  const byNode: Record<string, Record<string, string>> = {};
  for (const node of nodes) {
    const local: Record<string, string> = (byNode[node.id] = {});
    for (const email of node.emails) {
      for (const m of email.body.matchAll(EMIT_IN_BODY)) global[m[1]] = email.id;
      for (const n of Object.keys(email.answer?.emits ?? {})) global[n] = email.id;
      for (const op of email.answer?.ops ?? []) if (op.create) local[op.create] = email.id;
    }
  }
  return { global, byNode };
}

// Return `email` with any missing `date` edges its answer's anchor refs require. Add-only and
// idempotent (upgrades a `static` edge to `date`; never deletes), so it's safe to run on every edit.
export function withDerivedDateEdges(email: Email, nodeId: string, nodes: CorpusNode[]): Email {
  const refs = answerAnchorRefs(email.answer);
  if (!refs.size) return email;
  const { global, byNode } = emitters(nodes);
  const local = byNode[nodeId] ?? {};
  let deps = email.depends_on;
  for (const name of refs) {
    const src = local[name] ?? global[name];
    if (!src || src === email.id) continue;
    if (deps.some((d) => d.email === src && d.type === "date")) continue;
    const stat = deps.find((d) => d.email === src);
    deps = stat ? deps.map((d) => (d === stat ? { ...d, type: "date" } : d)) : [...deps, { email: src, type: "date" } as Edge];
  }
  return deps === email.depends_on ? email : { ...email, depends_on: deps };
}

// The set of nodes needed to validate `rootId` on its own: the storyline itself plus any other it
// actually reaches through a cross-node dependency — a depends_on edge into another node, a whole-node
// edge, or an answer @anchor whose only emitter lives elsewhere. Self-contained storylines (the norm)
// return just [root], so per-author (focus-mode) validation stays O(one scenario). Transitive: A->B->C
// pulls all three. Used to scope lint/oracle to one author's link without missing a real dependency.
export function focusClosure(nodes: CorpusNode[], rootId: string): CorpusNode[] {
  const byId = new Map(nodes.map((n) => [n.id, n]));
  const emailNode = new Map<string, string>();
  for (const n of nodes) for (const e of n.emails) emailNode.set(e.id, n.id);
  const { global } = emitters(nodes); // anchor name -> emitting email id (body {!=}/emits are corpus-wide)
  const seen = new Set<string>();
  const stack = [rootId];
  while (stack.length) {
    const id = stack.pop()!;
    if (seen.has(id) || !byId.has(id)) continue;
    seen.add(id);
    const push = (nid?: string) => { if (nid && !seen.has(nid)) stack.push(nid); };
    const fromEdge = (d: Edge) => { push(d.node); if (d.email) push(emailNode.get(d.email)); };
    const node = byId.get(id)!;
    for (const d of node.node_depends_on ?? []) fromEdge(d);
    for (const e of node.emails) {
      for (const d of e.depends_on) fromEdge(d);
      for (const name of answerAnchorRefs(e.answer)) push(emailNode.get(global[name] ?? ""));
    }
  }
  return nodes.filter((n) => seen.has(n.id));
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
