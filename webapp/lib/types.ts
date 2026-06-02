// TypeScript mirror of the on-disk corpus schema (sb/schema.py). These shapes
// serialize 1:1 to corpus/nodes/*.json — what sb.schema.load_corpus globs — so the
// app never invents a format of its own.
//
// The answer-key grammar types (Verb, ObjKind, EdgeType, Predicate, Op, Answer) are
// GENERATED from sb/schema.py during `npm run vendor` — see ./schema.generated.ts. They
// are re-exported here so the rest of the app imports everything from "@/lib/types" as
// before, but they cannot drift from the grammar the grader enforces. The structural types
// below (Edge, Email, CorpusNode) and the API result shapes are stable and hand-written.

export type { Verb, ObjKind, EdgeType, Predicate, Op, Answer } from "./schema.generated";
import type { EdgeType, Answer } from "./schema.generated";

export interface Edge {
  email?: string; // prerequisite email id
  node?: string; // OR a whole node (authoring sugar)
  type: EdgeType;
}

export type Tier = "T1" | "T2" | "T3";

export interface Email {
  id: string;
  from: string; // cast key
  to: string | string[]; // cast key(s)
  subject: string;
  body: string; // prose with {tokens}
  depends_on: Edge[];
  answer: Answer;
  tier?: Tier; // author-tagged difficulty, for score-by-tier
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
