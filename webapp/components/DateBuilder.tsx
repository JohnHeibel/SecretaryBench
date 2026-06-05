"use client";
import { useEffect, useMemo, useRef, useState } from "react";
import { resolveExpr } from "@/lib/api";
import { NTH, UNITS, WEEKDAYS } from "@/lib/grammar";
import {
  type Base, type BaseKind, type Expr, type MonthRef, type TimeHM, type TimeOfDay,
  addHour, canCarryTime, emptyBase, hhmm, humanTime, parseExpr, parseHHMM, serializeExpr, serveExpr, timeError,
} from "@/lib/dateExpr";

// The inline "fill-in-the-blank" date builder (design doc Direction A). An author composes a
// date from a short sentence of recognition-based choices — a starting point, then any offsets —
// instead of dragging Scratch blocks. ExprEditor edits the structured Expr (lib/dateExpr.ts)
// recursively (the `from` reference and `week of` inner are sub-editors on the SAME structure, so
// an in-progress pick is never lost to a string round-trip). DateBuilder wraps it for the string
// the answer key/body store, shows the day resolved by the REAL grader live, and offers a raw-text
// escape hatch (validated the same way) for anything the sentence can't express.

interface Props {
  value: string;                       // current grammar expression (predicate slot)
  anchors: string[];                   // @anchor names other emails publish
  serveDate: string;                   // preview "now"
  onChange: (expr: string) => void;
  allowTime?: boolean;                 // offer a clock suffix (events on an exact day); off for to-dos / deadlines
  anchorOrigins?: AnchorOrigins;       // name -> the email that published it, so the picker can show provenance
}

// name -> the email that published an anchor (subject for display). Threaded down to the anchor
// <select> so each option reads "@kickoff · from 'Project kickoff'" instead of a bare name.
export type AnchorOrigins = Record<string, { email: string; subject: string }>;
const anchorLabel = (name: string, origins?: AnchorOrigins) => { const o = origins?.[name]; return o?.subject ? `@${name} · ${o.subject}` : `@${name}`; };

// dropdown order/labels for the first blank — reads as the start of the sentence. The interval
// bases (week_of / month) are intentionally not offered here; they remain in the grammar and the
// raw "type it" escape hatch, just not in the common-case builder.
const BASE_OPTIONS: { kind: BaseKind; label: string }[] = [
  { kind: "serve", label: "the day this email arrives" },
  { kind: "anchor", label: "a date from an earlier email" },
  { kind: "next", label: "the next weekday (e.g. next Thursday)" },
  { kind: "this", label: "a weekday this week (e.g. this Friday)" },
  { kind: "nth", label: "the Nth weekday of a month (e.g. 3rd Friday)" },
  { kind: "dom", label: "a specific day of the month (e.g. the 25th)" },
];
const OFFERED_BASE = new Set<BaseKind>(BASE_OPTIONS.map((o) => o.kind));

const WD_LABEL: Record<string, string> = { MON: "Monday", TUE: "Tuesday", WED: "Wednesday", THU: "Thursday", FRI: "Friday", SAT: "Saturday", SUN: "Sunday" };
const NTH_LABEL: Record<string, string> = { "1": "1st", "2": "2nd", "3": "3rd", "4": "4th", "5": "5th", last: "last" };

const sel = "rounded border border-slate-700 bg-slate-800 px-1.5 py-1 text-xs text-slate-200 outline-none focus:border-sky-500";
const numField = "w-12 rounded border border-slate-700 bg-slate-800 px-1.5 py-1 text-xs text-slate-200 outline-none focus:border-sky-500";
const timeField = "rounded border border-slate-700 bg-slate-800 px-1 py-0.5 text-xs text-slate-200 outline-none focus:border-sky-500 [color-scheme:dark]";

// Is the expression still missing a required pick (so we shouldn't try to resolve it yet)?
function incomplete(e: Expr): boolean { return baseIncomplete(e.base); }
function baseIncomplete(b: Base): boolean {
  if (b.kind === "anchor") return !b.name;
  if (b.kind === "next" || b.kind === "this") return b.from !== null && incomplete(b.from);
  if (b.kind === "weekOf") return incomplete(b.inner);
  return false;
}
const hasAnchorRef = (s: string) => /@[A-Za-z_]/.test(s);

export default function DateBuilder({ value, anchors, serveDate, onChange, allowTime, anchorOrigins }: Props) {
  // Local structured state is the source of truth while editing; seeded from `value`. A non-empty
  // value that can't be represented (legacy / hand-typed exotic expr) opens in raw mode.
  const [expr, setExpr] = useState<Expr | null>(() => (value ? parseExpr(value) : null));
  const [raw, setRaw] = useState(value);
  const [mode, setMode] = useState<"builder" | "raw">(() => (value && parseExpr(value) === null ? "raw" : "builder"));
  const lastEmit = useRef(value);   // the last string WE emitted, so an external change is distinguishable from our own

  const serialized = mode === "raw" ? raw : expr ? serializeExpr(expr) : "";

  // Re-seed on an EXTERNAL value change (op switch, any_of reorder). We compare against our own last
  // emit (not the live serialization) so an in-progress incomplete pick — which we emit as "" — isn't
  // clobbered when that empty value comes back as the prop.
  useEffect(() => {
    if (value === lastEmit.current) return;
    const p = value ? parseExpr(value) : null;
    lastEmit.current = value;
    setRaw(value); setExpr(p); setMode(value && p === null ? "raw" : "builder");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [value]);

  const emit = (s: string) => { lastEmit.current = s; onChange(s); };
  // An incomplete pick (anchor with no name, an empty `from`, …) serializes to invalid grammar; emit ""
  // instead so the slot reads as simply unset (clean "add the date" guidance) rather than persisting "@".
  // Drop a stale clock if the base can no longer carry one (switched to a week/month, or added a `from`).
  const commit = (raw: Expr) => {
    const next = raw.time && !canCarryTime(raw) ? { ...raw, time: undefined } : raw;
    setExpr(next); setRaw(serializeExpr(next)); emit(incomplete(next) ? "" : serializeExpr(next));
  };
  const showTime = allowTime !== false && canCarryTime(expr);

  return (
    <div className="space-y-1">
      {mode === "raw" ? (
        <RawField value={raw} onChange={(v) => { setRaw(v); emit(v); setExpr(parseExpr(v)); }}
          onUseBuilder={() => { setMode("builder"); setExpr(parseExpr(raw)); }} />
      ) : (
        <div className="flex flex-wrap items-center gap-1.5">
          <ExprEditor expr={expr} onChange={commit} anchors={anchors} serveDate={serveDate} anchorOrigins={anchorOrigins} />
          {showTime && expr && <TimeControl time={expr.time ?? null} onChange={(time) => commit({ ...expr, time })} />}
          <button type="button" onClick={() => { setRaw(serialized); setMode("raw"); }}
            className="ml-auto text-[11px] text-slate-500 hover:text-sky-300" title="Type the expression yourself. It is still checked live.">type it</button>
        </div>
      )}
      {mode === "builder" && (
        <p className="text-[11px] leading-snug text-slate-500">Pick a starting day, then optionally shift it forward or back, like &ldquo;+2 weeks&rdquo;. Think &ldquo;two weeks after the kickoff&rdquo;. The app fills in the real date.</p>
      )}
      <Preview serialized={serialized} serveDate={serveDate}
        incomplete={mode === "builder" && (!expr || incomplete(expr))}
        time={mode === "builder" ? expr?.time ?? null : null} />
    </div>
  );
}

// The optional clock suffix. Collapsed to a "+ add a time" affordance by default (many dates are
// day-level); when set, a start and an optional end (a within-day span), with the resolver's
// work-hours rule surfaced live via timeError — important because the live preview skips resolving
// anchor-based exprs, so this is the only place an out-of-hours time on a needle gets flagged.
function TimeControl({ time, onChange }: { time: TimeOfDay | null; onChange: (t: TimeOfDay | undefined) => void }) {
  if (!time) return (
    <button type="button" onClick={() => onChange({ start: { h: 9, m: 0 }, end: { h: 10, m: 0 } })}
      className="text-[11px] text-slate-500 hover:text-sky-300" title="Give this event a clock time (work hours 5 AM–11 PM).">+ add a time</button>
  );
  const err = timeError(time);
  const setStart = (v: TimeHM | null) => v && onChange({ ...time, start: v });
  const setEnd = (v: TimeHM | null) => v && onChange({ ...time, end: v });
  return (
    <span className="flex flex-wrap items-center gap-1 rounded border border-slate-700/70 bg-slate-900/60 px-1.5 py-1">
      <span className="text-[11px] text-slate-500">at</span>
      <input type="time" step={60} min="05:00" max="23:00" value={hhmm(time.start)} onChange={(e) => setStart(parseHHMM(e.target.value))} className={timeField} />
      {time.end ? (
        <>
          <span className="text-slate-500">–</span>
          <input type="time" step={60} min="05:00" max="23:00" value={hhmm(time.end)} onChange={(e) => setEnd(parseHHMM(e.target.value))} className={timeField} />
          <button type="button" onClick={() => onChange({ ...time, end: null })} title="make it a single start time (no end)" className="px-0.5 text-[11px] text-slate-500 hover:text-sky-300">·pt</button>
        </>
      ) : (
        <button type="button" onClick={() => onChange({ ...time, end: addHour(time.start) })} className="text-[11px] text-slate-500 hover:text-sky-300">+ end</button>
      )}
      <button type="button" onClick={() => onChange(undefined)} title="remove the time (grade at day level)" className="px-0.5 text-slate-500 hover:text-rose-400">✕</button>
      {err && <span className="ml-0.5 text-[11px] text-amber-400" title="work hours are 5 AM–11 PM">{err}</span>}
    </span>
  );
}

// --- the recursive structured editor (no strings) --------------------------

function ExprEditor({ expr, onChange, anchors, serveDate, anchorOrigins }: { expr: Expr | null; onChange: (e: Expr) => void; anchors: string[]; serveDate: string; anchorOrigins?: AnchorOrigins }) {
  const setBaseKind = (kind: BaseKind) => {
    const keep = kind !== "weekOf" && kind !== "month";   // interval bases can't carry offsets or a clock
    onChange({ base: emptyBase(kind), offsets: keep && expr ? expr.offsets : [], time: keep ? expr?.time : undefined });
  };
  const isIntervalBase = expr && (expr.base.kind === "weekOf" || expr.base.kind === "month");
  return (
    <span className="flex flex-wrap items-center gap-1.5 text-xs text-slate-300">
      <select value={expr?.base.kind ?? ""} onChange={(e) => setBaseKind(e.target.value as BaseKind)} className={sel}>
        <option value="" disabled>start from…</option>
        {BASE_OPTIONS.map((o) => <option key={o.kind} value={o.kind}>{o.label}</option>)}
        {expr && !OFFERED_BASE.has(expr.base.kind) && <option value={expr.base.kind}>advanced: {expr.base.kind}</option>}
      </select>
      {expr && <BaseControls base={expr.base} anchors={anchors} serveDate={serveDate} anchorOrigins={anchorOrigins} onChange={(b) => onChange({ ...expr, base: b })} />}
      {expr && !isIntervalBase && <OffsetRows offsets={expr.offsets} onChange={(offsets) => onChange({ ...expr, offsets })} />}
    </span>
  );
}

function BaseControls({ base, anchors, serveDate, anchorOrigins, onChange }: { base: Base; anchors: string[]; serveDate: string; anchorOrigins?: AnchorOrigins; onChange: (b: Base) => void }) {
  switch (base.kind) {
    case "serve":
      return null;
    case "anchor":
      return (
        <select value={base.name} onChange={(e) => onChange({ kind: "anchor", name: e.target.value })} className={sel}>
          <option value="">pick a published date…</option>
          {anchors.map((a) => <option key={a} value={a}>{anchorLabel(a, anchorOrigins)}</option>)}
          {base.name && !anchors.includes(base.name) && <option value={base.name}>{anchorLabel(base.name, anchorOrigins)}</option>}
        </select>
      );
    case "next":
    case "this": {
      const b = base;
      return (
        <span className="flex flex-wrap items-center gap-1.5">
          <WeekdaySelect wd={b.wd} onChange={(wd) => onChange({ ...b, wd })} />
          {b.from === null ? (
            <button type="button" onClick={() => onChange({ ...b, from: serveExpr() })}
              className="text-[11px] text-slate-500 hover:text-sky-300">+ counting from another date</button>
          ) : (
            <span className="flex flex-wrap items-center gap-1.5 rounded border border-slate-700/70 bg-slate-900/60 px-1.5 py-1">
              <span className="text-[11px] text-slate-500">from</span>
              <ExprEditor expr={b.from} anchors={anchors} serveDate={serveDate} anchorOrigins={anchorOrigins} onChange={(from) => onChange({ ...b, from })} />
              <button type="button" onClick={() => onChange({ ...b, from: null })} className="text-[11px] text-slate-500 hover:text-rose-400">✕</button>
            </span>
          )}
        </span>
      );
    }
    case "nth":
      return (
        <span className="flex items-center gap-1.5">
          <select value={base.n} onChange={(e) => onChange({ ...base, n: e.target.value })} className={sel}>
            {NTH.map((n) => <option key={n} value={n}>{NTH_LABEL[n]}</option>)}
          </select>
          <WeekdaySelect wd={base.wd} onChange={(wd) => onChange({ ...base, wd })} />
          <span className="text-slate-500">of</span>
          <MonthRefSelect month={base.month} onChange={(month) => onChange({ ...base, month })} />
        </span>
      );
    case "dom":
      return (
        <span className="flex items-center gap-1.5">
          <input type="number" min={1} max={31} value={base.day} onChange={(e) => onChange({ ...base, day: clampInt(e.target.value, 1, 31) })} className={numField} />
          <span className="text-slate-500">of</span>
          <MonthRefSelect month={base.month} onChange={(month) => onChange({ ...base, month })} />
        </span>
      );
    case "weekOf":
      return (
        <span className="flex flex-wrap items-center gap-1.5 rounded border border-slate-700/70 bg-slate-900/60 px-1.5 py-1">
          <ExprEditor expr={base.inner} anchors={anchors} serveDate={serveDate} anchorOrigins={anchorOrigins} onChange={(inner) => onChange({ kind: "weekOf", inner })} />
        </span>
      );
    case "month":
      return <MonthRefSelect month={base.month} onChange={(month) => onChange({ kind: "month", month })} />;
  }
}

function WeekdaySelect({ wd, onChange }: { wd: string; onChange: (wd: string) => void }) {
  return (
    <select value={wd} onChange={(e) => onChange(e.target.value)} className={sel}>
      {WEEKDAYS.map((d) => <option key={d} value={d}>{WD_LABEL[d]}</option>)}
    </select>
  );
}

function MonthRefSelect({ month, onChange }: { month: MonthRef; onChange: (m: MonthRef) => void }) {
  return (
    <span className="flex items-center gap-1.5">
      <select value={month.sign} onChange={(e) => onChange(e.target.value === "0" ? { sign: "0", n: 0 } : { sign: e.target.value as "+" | "-", n: month.n || 1 })} className={sel}>
        <option value="0">this month</option>
        <option value="+">months ahead</option>
        <option value="-">months back</option>
      </select>
      {month.sign !== "0" && (
        <input type="number" min={1} value={month.n} onChange={(e) => onChange({ ...month, n: clampInt(e.target.value, 1, 120) })} className={numField} />
      )}
    </span>
  );
}

// A date "shift" is plain calendar arithmetic on the starting day (e.g. "+2 weeks" = two weeks
// after it). We never say "offset" to authors — the row reads "shift by [2] [weeks] [later]" and
// the sign is a friendly later/earlier picker over the same +/- grammar.
function OffsetRows({ offsets, onChange }: { offsets: Expr["offsets"]; onChange: (o: Expr["offsets"]) => void }) {
  const set = (i: number, patch: Partial<Expr["offsets"][number]>) => onChange(offsets.map((o, j) => (j === i ? { ...o, ...patch } : o)));
  return (
    <span className="flex flex-wrap items-center gap-1.5">
      {offsets.map((o, i) => (
        <span key={i} className="flex items-center gap-1 rounded border border-slate-700/70 bg-slate-900/60 px-1 py-0.5">
          <span className="text-[11px] text-slate-500">shift by</span>
          <input type="number" min={0} value={o.n} onChange={(e) => set(i, { n: clampInt(e.target.value, 0, 9999) })} className={numField} />
          <select value={o.unit} onChange={(e) => set(i, { unit: e.target.value as Expr["offsets"][number]["unit"] })} className={sel}>
            {UNITS.map((u) => <option key={u.value} value={u.value}>{u.label}</option>)}
          </select>
          <select value={o.sign} onChange={(e) => set(i, { sign: e.target.value as "+" | "-" })} className={sel} title="later = after the starting day; earlier = before it">
            <option value="+">later</option>
            <option value="-">earlier</option>
          </select>
          <button type="button" onClick={() => onChange(offsets.filter((_, j) => j !== i))} className="px-0.5 text-slate-500 hover:text-rose-400">✕</button>
        </span>
      ))}
      <button type="button" onClick={() => onChange([...offsets, { sign: "+", n: 1, unit: "d" }])}
        className="text-[11px] text-slate-500 hover:text-sky-300" title="Shift the starting day forward or back, like +2 weeks.">+ shift the date (e.g. +1 week)</button>
    </span>
  );
}

// --- raw escape hatch + live preview ---------------------------------------

function RawField({ value, onChange, onUseBuilder }: { value: string; onChange: (v: string) => void; onUseBuilder: () => void }) {
  return (
    <div className="flex items-center gap-2">
      <input value={value} onChange={(e) => onChange(e.target.value)} spellCheck={false}
        placeholder="e.g. @signing+2w, next:THU, week_of:(serve+1w)"
        className="min-w-[16rem] flex-1 rounded border border-slate-700 bg-slate-800 px-2 py-1 font-mono text-xs text-sky-300 outline-none focus:border-sky-500" />
      <button type="button" onClick={onUseBuilder} className="text-[11px] text-slate-500 hover:text-sky-300">use builder</button>
    </div>
  );
}

function Preview({ serialized, serveDate, incomplete, time }: { serialized: string; serveDate: string; incomplete: boolean; time: TimeOfDay | null }) {
  const [res, setRes] = useState<{ ok: boolean; text: string } | null>(null);
  const anchorRef = useMemo(() => hasAnchorRef(serialized), [serialized]);

  useEffect(() => {
    if (incomplete || !serialized.trim() || anchorRef) { setRes(null); return; }
    let alive = true;
    const t = setTimeout(async () => {
      const r = await resolveExpr(serialized, serveDate, {});
      if (alive) setRes({ ok: !!r.ok, text: r.ok ? (r.human ?? "") : (r.error ?? "invalid") });
    }, 200);
    return () => { alive = false; clearTimeout(t); };
  }, [serialized, serveDate, incomplete, anchorRef]);

  const timeWarn = time ? timeError(time) : null;     // client-side, so it shows even on anchor exprs

  return (
    <div className="min-h-[1rem] text-[11px] leading-snug">
      {incomplete || !serialized.trim() ? <span className="text-slate-500">build a date…</span>
        : anchorRef ? <span className="text-amber-400">→ uses a published date{time ? ` (${humanTime(time)})` : ""}; resolves when this email is sent</span>
        : res ? <span className={res.ok ? "text-emerald-400" : "text-rose-400"}>{res.ok ? "→ " : "⚠ "}{res.text}</span>
        : <span className="text-slate-500">…</span>}
      {timeWarn ? <span className="ml-2 text-amber-400">⚠ {timeWarn}</span> : null}
    </div>
  );
}

function clampInt(v: string, lo: number, hi: number): number {
  const n = parseInt(v, 10);
  if (Number.isNaN(n)) return lo;
  return Math.min(hi, Math.max(lo, n));
}
