// AUTO-GENERATED from sb/schema.py by scripts/vendor_sb.py — DO NOT EDIT BY HAND.
// Regenerated on every `npm run vendor` / `npm run build`. To change the answer-key grammar,
// edit the Python grammar lists (VERB_ORDER / KIND_ORDER / EDGE_ORDER / PREDICATE_OPS) and
// rebuild; the editor then fails typecheck if it still uses a field the grammar dropped.

export type Verb = "create" | "move" | "cancel";
export type ObjKind = "event" | "todo";
export type EdgeType = "static" | "date";

// Date predicate over an answer-key slot. Exactly one key is set; any_of takes a list.
export interface Predicate {
  eq?: string;
  by?: string;
  in?: string;
  not_in?: string;
  any_of?: string[];
}

// One verb on a named obligation. Serializes 1:1 to corpus JSON: the verb is the KEY and its
// value is the obligation name, e.g. { "create": "kickoff", "kind": "event", "on": {...} }.
export interface Op {
  create?: string;
  move?: string;
  cancel?: string;
  kind?: ObjKind;
  on?: Predicate;
  match?: string[];
  tolerance?: string;
}

export interface Answer {
  ops: Op[];
  emits?: Record<string, string>;
}
