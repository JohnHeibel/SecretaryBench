// TypeScript mirror of the on-disk corpus schema (sb/schema.py). These shapes
// serialize 1:1 to corpus/nodes/*.json — what sb.schema.load_corpus globs — so the
// app never invents a format of its own.

export type EdgeType = "static" | "date";

export interface Edge {
  email?: string; // prerequisite email id
  node?: string; // OR a whole node (authoring sugar)
  type: EdgeType;
}

// Answer-key predicate over a date slot. Exactly one key is set.
export interface Predicate {
  eq?: string;
  in?: string;
  by?: string;
  any_of?: string[];
  not_in?: string;
}

export interface ExpectEntry {
  action: "create_event" | "create_todo" | "reschedule" | "reply" | "delegate";
  title_match?: string[];
  start?: Predicate; // events
  due?: Predicate; // todos
  count?: number;
  tolerance?: string; // "exact_day" | "within:Nd"
  recurrence?: Record<string, unknown>;
}

export interface ForbidEntry {
  action?: string; // omitted = any action
  title_match?: string[];
}

export interface Answer {
  expect: ExpectEntry[];
  forbid?: ForbidEntry[];
  emits?: Record<string, string>; // anchor name -> expr (non-rendered)
}

export interface Email {
  id: string;
  from: string; // cast key
  to: string | string[]; // cast key(s)
  subject: string;
  body: string; // prose with {tokens}
  depends_on: Edge[];
  answer: Answer;
}

export interface CorpusNode {
  id: string;
  cast: Record<string, string>; // key -> display name
  node_depends_on?: Edge[];
  emails: Email[];
}

// --- validation API shapes (Python /api/lint, /api/resolve) ---

export interface ResolveResult {
  ok: boolean;
  kind?: "date" | "datetime" | "interval";
  iso?: string;
  human?: string;
  error?: string;
}

export interface LintEmail {
  id: string;
  node: string;
  depends_on: Edge[];
  emits: string[];
  anchor_refs: string[];
}

export interface LintResult {
  ok: boolean;
  error?: string;
  summary?: { nodes: number; emails: number; anchors: number };
  emission_map?: Record<string, string>; // anchor -> email id
  emails?: LintEmail[];
  order?: string[];
}

// /api/oracle — satisfiability: did the reference solver score 1.0? `failures`
// lists email ids whose answer key the perfect secretary could NOT satisfy.
export interface OracleResult {
  ok: boolean;
  error?: string;
  score?: number;
  passed?: number;
  total?: number;
  failures?: string[];
}
