// A structured mirror of the date grammar (sb/resolver.py / ANSWER_KEY_GRAMMAR.md §3),
// with a parser and serializer. This powers the inline "fill-in-the-blank" date builder:
// the UI edits a structured Expr, serializes it to the grammar string the answer key/body
// store, and the REAL Python resolver (via /api/resolve) remains the only source of truth
// for what a string actually evaluates to. parse/serialize must stay faithful to the Python
// parser — every production it accepts, this accepts; anything it can't represent, parseExpr
// returns null so the UI falls back to the raw-expression escape hatch.
//
//   expr      := base (offset)* (time)?      ( base defaults to `serve` if the string starts with an offset )
//   base      := serve | @NAME | next:WD [from expr] | this:WD [from expr]
//              | nth:(N|last),WD,monthref | dom:D,monthref | week_of:(expr) | month:monthref
//   offset    := (+|-) INT (d|bd|w|m|y)
//   monthref  := 0m | (+|-) INT m
//   time      := @ HH:MM [ - HH:MM ]         ( static clock; work hours 05:00–23:00, end > start )

export type Unit = "d" | "bd" | "w" | "m" | "y";
export interface Offset { sign: "+" | "-"; n: number; unit: Unit; }
export interface MonthRef { sign: "0" | "+" | "-"; n: number; }  // "0" => serve's month (n ignored)

// An optional clock suffix on a date: `@HH:MM` (a point) or `@HH:MM-HH:MM` (a within-day span).
// Static — it rides inside the one governed expr (so body + answer can't drift) but does not shift
// with serve order. The resolver bounds it by the CEO's work hours; timeError() mirrors that check.
export interface TimeHM { h: number; m: number; }
export interface TimeOfDay { start: TimeHM; end: TimeHM | null; }   // end null => a point time

export const WORK_START_MIN = 5 * 60;    // 05:00 — mirror sb/resolver.py WORK_START
export const WORK_END_MIN = 23 * 60;     // 23:00 — mirror sb/resolver.py WORK_END
const mins = (t: TimeHM) => t.h * 60 + t.m;

export type Base =
  | { kind: "serve" }
  | { kind: "anchor"; name: string }
  | { kind: "next"; wd: string; from: Expr | null }    // from null => serve-relative
  | { kind: "this"; wd: string; from: Expr | null }
  | { kind: "nth"; n: string; wd: string; month: MonthRef }   // n: "1".."5" | "last"
  | { kind: "dom"; day: number; month: MonthRef }
  | { kind: "weekOf"; inner: Expr }
  | { kind: "month"; month: MonthRef };

export interface Expr { base: Base; offsets: Offset[]; time?: TimeOfDay | null; }

export const BASE_KINDS = ["serve", "anchor", "next", "this", "nth", "dom", "weekOf", "month"] as const;
export type BaseKind = (typeof BASE_KINDS)[number];

// --- constructors / defaults ----------------------------------------------

export const serveExpr = (): Expr => ({ base: { kind: "serve" }, offsets: [] });
export const thisMonth = (): MonthRef => ({ sign: "0", n: 0 });

// A blank base of the given kind, with sensible defaults so a freshly-picked base is valid-ish.
export function emptyBase(kind: BaseKind): Base {
  switch (kind) {
    case "serve": return { kind: "serve" };
    case "anchor": return { kind: "anchor", name: "" };
    case "next": return { kind: "next", wd: "MON", from: null };
    case "this": return { kind: "this", wd: "MON", from: null };
    case "nth": return { kind: "nth", n: "1", wd: "MON", month: thisMonth() };
    case "dom": return { kind: "dom", day: 1, month: thisMonth() };
    case "weekOf": return { kind: "weekOf", inner: serveExpr() };
    case "month": return { kind: "month", month: thisMonth() };
  }
}

// An interval-typed expr (week_of / whole month with no day-collapsing offset) — what the
// `within` / `not within` predicates need. Any offset collapses an interval to a date in the
// resolver, so an interval is base-only.
export function isInterval(e: Expr): boolean {
  if (e.time) return false;                                  // a timed day is a clock span, not a week/month
  return (e.base.kind === "weekOf" || e.base.kind === "month") && e.offsets.length === 0;
}

// --- time-of-day helpers (shared with the builder; pure) -------------------

// Can a top-level clock suffix attach to this expr at all? `week_of`/`month` are interval bases
// (a clock on a whole week/month is meaningless), and a `next/this … from <ref>` clause greedily
// consumes the rest of the string in the grammar, so a top-level time can't follow it.
export function canCarryTime(e: Expr | null): boolean {
  if (!e) return false;
  const b = e.base;
  if (b.kind === "weekOf" || b.kind === "month") return false;
  if ((b.kind === "next" || b.kind === "this") && b.from !== null) return false;
  return true;
}

// Mirror sb/resolver.py _check_work_hours so the builder flags an unbookable time live (even for
// anchor-based exprs the live resolver preview skips). Returns a human reason, or null when valid.
export function timeError(t: TimeOfDay): string | null {
  if (mins(t.start) < WORK_START_MIN || mins(t.start) > WORK_END_MIN) return "start must be 05:00–23:00 (work hours)";
  if (t.end) {
    if (mins(t.end) <= mins(t.start)) return "end must be after the start";
    if (mins(t.end) > WORK_END_MIN) return "end must be by 23:00";
  }
  return null;
}

export const hhmm = (t: TimeHM): string => `${String(t.h).padStart(2, "0")}:${String(t.m).padStart(2, "0")}`;
export function parseHHMM(s: string): TimeHM | null {
  const m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  const h = parseInt(m[1], 10), mm = parseInt(m[2], 10);
  return h >= 0 && h <= 23 && mm >= 0 && mm <= 59 ? { h, m: mm } : null;
}
export const addHour = (t: TimeHM): TimeHM => ({ h: Math.min(23, t.h + 1), m: t.m });

// Natural-language clock, matching sb/resolver.py human() ("2–3 PM", "9 AM").
export function humanTime(t: TimeOfDay): string {
  const one = (x: TimeHM) => `${x.h % 12 || 12}${x.m ? `:${String(x.m).padStart(2, "0")}` : ""} ${x.h < 12 ? "AM" : "PM"}`;
  return t.end ? `${one(t.start)}–${one(t.end)}` : one(t.start);
}

// --- serialize (structured -> grammar string) ------------------------------

function serializeMonthRef(m: MonthRef): string {
  return m.sign === "0" ? "0m" : `${m.sign}${m.n}m`;
}

function serializeBase(b: Base): string {
  switch (b.kind) {
    case "serve": return "serve";
    case "anchor": return `@${b.name}`;
    case "next":
    case "this": return `${b.kind}:${b.wd}${b.from ? ` from ${serializeExpr(b.from)}` : ""}`;
    case "nth": return `nth:${b.n},${b.wd},${serializeMonthRef(b.month)}`;
    case "dom": return `dom:${b.day},${serializeMonthRef(b.month)}`;
    case "weekOf": return `week_of:(${serializeExpr(b.inner)})`;
    case "month": return `month:${serializeMonthRef(b.month)}`;
  }
}

function serializeTime(t: TimeOfDay): string {
  return t.end ? `@${hhmm(t.start)}-${hhmm(t.end)}` : `@${hhmm(t.start)}`;
}

export function serializeExpr(e: Expr): string {
  const core = e.offsets.reduce((s, o) => `${s}${o.sign}${o.n}${o.unit}`, serializeBase(e.base));
  return e.time ? `${core} ${serializeTime(e.time)}` : core;
}

// --- parse (grammar string -> structured) ----------------------------------
// Mirrors resolver._parse_expr / _parse_base. Returns null on anything it can't represent,
// so the caller drops to the raw-text escape hatch instead of silently dropping detail.

const OFFSET_RE = /^\s*([+-]\d+)(bd|d|w|m|y)/;
const FROM_RE = /^\s+from\s+/i;
const TIME_RE = /^\s*@(\d{1,2}):(\d{2})(?:\s*-\s*(\d{1,2}):(\d{2}))?/;   // mirrors resolver _TIME_RE

function monthRefFromInt(numStr: string): MonthRef {
  const n = parseInt(numStr, 10);
  if (n === 0) return { sign: "0", n: 0 };
  return n > 0 ? { sign: "+", n } : { sign: "-", n: -n };
}

// next:/this: optional ` from <expr>` — the from-clause consumes the rest of the string as a
// full sub-expression (resolver._parse_from_clause), so nothing follows it.
function parseFromClause(rest: string): { from: Expr | null; rest: string; ok: boolean } {
  const m = rest.match(FROM_RE);
  if (!m) return { from: null, rest, ok: true };          // no clause: serve-relative
  const inner = parseExpr(rest.slice(m[0].length));
  return { from: inner, rest: "", ok: inner !== null };   // clause present but unparseable => not ok
}

function parseBase(input: string): { base: Base; rest: string } | null {
  const s = input.replace(/^\s+/, "");

  if (s.startsWith("serve")) return { base: { kind: "serve" }, rest: s.slice(5) };

  let m = s.match(/^@([A-Za-z_][A-Za-z0-9_]*)/);
  if (m) return { base: { kind: "anchor", name: m[1] }, rest: s.slice(m[0].length) };

  m = s.match(/^(next|this):([A-Za-z]{3})/i);
  if (m) {
    const kind = m[1].toLowerCase() as "next" | "this";
    const fc = parseFromClause(s.slice(m[0].length));
    if (!fc.ok) return null;
    return { base: { kind, wd: m[2].toUpperCase(), from: fc.from }, rest: fc.rest };
  }

  m = s.match(/^nth:(\d+|last),([A-Za-z]{3}),([+-]?\d+)m/i);
  if (m) {
    const n = m[1].toLowerCase() === "last" ? "last" : String(parseInt(m[1], 10));
    return { base: { kind: "nth", n, wd: m[2].toUpperCase(), month: monthRefFromInt(m[3]) }, rest: s.slice(m[0].length) };
  }

  m = s.match(/^dom:(\d+),([+-]?\d+)m/i);
  if (m) return { base: { kind: "dom", day: parseInt(m[1], 10), month: monthRefFromInt(m[2]) }, rest: s.slice(m[0].length) };

  m = s.match(/^month:([+-]?\d+)m/i);
  if (m) return { base: { kind: "month", month: monthRefFromInt(m[1]) }, rest: s.slice(m[0].length) };

  if (/^week_of:/i.test(s)) {
    let rest = s.slice("week_of:".length).replace(/^\s+/, "");
    if (!rest.startsWith("(")) return null;
    let depth = 0, i = 0;
    for (; i < rest.length; i++) {
      if (rest[i] === "(") depth++;
      else if (rest[i] === ")" && --depth === 0) break;
    }
    if (depth !== 0) return null;                          // unbalanced
    const inner = parseExpr(rest.slice(1, i));
    if (inner === null) return null;
    return { base: { kind: "weekOf", inner }, rest: rest.slice(i + 1) };
  }

  return null;
}

export function parseExpr(raw: string): Expr | null {
  const s = raw.trim();
  if (!s) return null;

  let base: Base, rest: string;
  if (OFFSET_RE.test(s)) { base = { kind: "serve" }; rest = s; }   // bare offset => serve-relative
  else {
    const pb = parseBase(s);
    if (!pb) return null;
    base = pb.base; rest = pb.rest;
  }

  const offsets: Offset[] = [];
  for (let m = rest.match(OFFSET_RE); m; m = rest.match(OFFSET_RE)) {
    offsets.push({ sign: m[1][0] as "+" | "-", n: Math.abs(parseInt(m[1], 10)), unit: m[2] as Unit });
    rest = rest.slice(m[0].length);
  }

  // optional trailing time-of-day suffix (@HH:MM or @HH:MM-HH:MM). We accept any structurally valid
  // clock (00:00–23:59) here; work-hours / end>start are VALUE checks the resolver (and timeError)
  // flag live — same split as dom:99 parsing through but failing at resolve.
  let time: TimeOfDay | undefined;
  const tm = rest.match(TIME_RE);
  if (tm) {
    const start = parseHHMM(`${tm[1]}:${tm[2]}`);
    if (!start) return null;
    let end: TimeHM | null = null;
    if (tm[3] !== undefined) { end = parseHHMM(`${tm[3]}:${tm[4]}`); if (!end) return null; }
    time = { start, end };
    rest = rest.slice(tm[0].length);
  }

  if (rest.trim()) return null;                            // trailing junk => can't represent
  return time ? { base, offsets, time } : { base, offsets };
}
