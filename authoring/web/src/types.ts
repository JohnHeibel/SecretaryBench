// Shapes mirrored from authoring/corpus_io.py + authoring/server.py.

export type BaseKind =
  | "serve"
  | "anchor"
  | "next_weekday"
  | "this_weekday"
  | "nth_weekday"
  | "day_of_month"
  | "month"
  | "week_of"
  | "raw";

export interface ChipBase {
  kind: BaseKind;
  name?: string; // anchor
  weekday?: string; // next/this/nth weekday (MON..SUN)
  n?: number | "last"; // nth
  month_offset?: number; // nth / day_of_month / month
  day?: number; // day_of_month
  inner?: Chip; // week_of
  token?: string; // raw
}

export interface Chip {
  base: ChipBase;
  offset?: { amount: number; unit: string } | null;
  time?: { hour: number; minute: number } | null;
  emit_as?: string | null;
}

export type Segment =
  | { type: "text"; value: string }
  | { type: "chip"; chip: Chip; token?: string }
  | { type: "fact"; name: string; value: string | null; token?: string };

export type MatchKind = "on" | "by" | "within" | "any_of";

export interface Predicate {
  match: MatchKind;
  chip?: Chip;
  chips?: Chip[]; // any_of
  avoid_chip?: Chip; // within + not_in
}

export type Action = "create_event" | "create_todo" | "reschedule";

export interface ExpectForm {
  action: Action;
  title_match: string[];
  when: Predicate | null;
  duration?: string | null; // literal "90m" OR a fact ref "@client_meeting_len"
  count?: number | string | null;
  tolerance?: string;
}

export interface ForbidForm {
  action?: string | null;
  title_match: string[];
}

export interface AnswerForm {
  expect: ExpectForm[];
  forbid: ForbidForm[];
  emits: Record<string, Chip>;
  facts?: Record<string, string>;
}

export interface Edge {
  email?: string;
  node?: string;
  type: "static" | "date";
  recommended?: "static" | "date";
}

export interface EmailForm {
  id: string;
  thread: string;
  from: string;
  to: string[];
  subject: string;
  body_segments: Segment[];
  depends_on: Edge[];
  answer: AnswerForm;
  emits: Record<string, Chip>;
  reachable_anchors: string[];
  defined_facts: Record<string, string>;
  reachable_facts: string[];
  // client-only layout (persisted in localStorage, not corpus)
  x?: number;
  y?: number;
}

export interface ThreadForm {
  id: string;
  cast: Record<string, string>;
  scenario: string;
  node_depends_on: Edge[];
  emails: string[];
}

export const DEFAULT_SCENARIO = "unsorted";

export interface Sample {
  ok: boolean;
  serve_date: Record<string, string>;
  anchors: Record<string, string>;
}

export interface Graph {
  threads: ThreadForm[];
  emails: Record<string, EmailForm>;
  emission_map: Record<string, string>;
  fact_map: Record<string, string>;
  fact_values: Record<string, string>;
  errors: string[];
  sample: Sample;
  start: string;
}

export interface ResolveResult {
  ok: boolean;
  kind?: string;
  human?: string;
  iso?: string;
  error?: string;
}

export interface OracleResult {
  ok: boolean;
  error?: string;
  score?: number;
  passed?: number;
  total?: number;
  results: Record<string, { passed: boolean; headline: string }>;
}
