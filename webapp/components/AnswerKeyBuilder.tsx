"use client";
import { useEffect, useState } from "react";
import { resolveExpr } from "@/lib/api";
import { OP_CHOICES, opChoiceId, opName, opVerb } from "@/lib/grammar";
import type { Answer, Op, Predicate } from "@/lib/types";
import TokenBlockly from "./blockly/TokenBlockly";

interface Props {
  answer: Answer;
  anchors: string[];
  serveDate: string;
  onChange: (answer: Answer) => void;
}

// How forgiving the day match is. exact_day is the default (and serializes to nothing);
// within:Nd lets the model's date land within N days of the answer.
const TOLERANCE_CHOICES: { value: string; label: string }[] = [
  { value: "exact_day", label: "the exact day" },
  { value: "within:1d", label: "within 1 day" },
  { value: "within:2d", label: "within 2 days" },
  { value: "within:3d", label: "within 3 days" },
  { value: "within:7d", label: "within a week" },
];

const PRED_OPS: { key: keyof Predicate; label: string }[] = [
  { key: "eq", label: "on exactly (eq)" },
  { key: "by", label: "on or before (by)" },
  { key: "in", label: "within interval (in)" },
  { key: "any_of", label: "any of (any_of)" },
  { key: "not_in", label: "not within (not_in)" },
];

// Live, debounced preview of a typed date expression — mirrors BodyEditor's chips.
// A comma-joined string (the any_of case) is previewed part-by-part. @anchor exprs
// resolve at serve time, so we don't try to resolve them here.
function DatePreview({ expr, serveDate }: { expr: string; serveDate: string }) {
  const [chips, setChips] = useState<{ expr: string; ok: boolean; text: string }[]>([]);
  useEffect(() => {
    const parts = expr.split(",").map((s) => s.trim()).filter(Boolean);
    if (!parts.length) { setChips([]); return; }
    let alive = true;
    const t = setTimeout(async () => {
      const out: { expr: string; ok: boolean; text: string }[] = [];
      for (const p of parts) {
        if (/@\w/.test(p)) { out.push({ expr: p, ok: true, text: "resolves at serve time" }); continue; }
        const r = await resolveExpr(p, serveDate, {});
        out.push({ expr: p, ok: !!r.ok, text: r.ok ? (r.human ?? "") : (r.error ?? "invalid") });
      }
      if (alive) setChips(out);
    }, 250);
    return () => { alive = false; clearTimeout(t); };
  }, [expr, serveDate]);
  if (!chips.length) return null;
  return (
    <div className="flex flex-col gap-0.5 text-[11px] leading-snug">
      {chips.map((c, k) => <span key={k} className={c.ok ? "text-emerald-400" : "text-rose-400"}>{c.ok ? "→" : "⚠"} {c.text}</span>)}
    </div>
  );
}

function predOp(p?: Predicate): keyof Predicate {
  if (!p) return "eq";
  return (Object.keys(p)[0] as keyof Predicate) ?? "eq";
}
function predExpr(p?: Predicate): string {
  if (!p) return "";
  const v = Object.values(p)[0];
  return Array.isArray(v) ? v.join(", ") : (v ?? "");
}

export default function AnswerKeyBuilder({ answer, anchors, serveDate, onChange }: Props) {
  // which op index is currently picking a date via the block builder
  const [picking, setPicking] = useState<number | null>(null);
  // which op indices have the (advanced) keyword override revealed
  const [showMatch, setShowMatch] = useState<Record<number, boolean>>({});
  const ops = answer.ops ?? [];
  const noAction = ops.length === 0;

  function setOps(next: Op[]) { onChange({ ...answer, ops: next }); }
  function patch(i: number, p: Partial<Op>) { setOps(ops.map((o, j) => j === i ? { ...o, ...p } : o)); }
  function replaceAt(i: number, op: Op) { setOps(ops.map((o, j) => j === i ? op : o)); }

  // Changing the verb means the JSON KEY changes (create -> move), so rebuild the op from
  // scratch, carrying the name/match/date and dropping whatever the new verb can't hold.
  function setChoice(i: number, choiceId: string) {
    const c = OP_CHOICES.find((x) => x.id === choiceId);
    if (!c) return;
    const cur = ops[i];
    const next: Op = { [c.verb]: opName(cur) };
    if (c.kind) next.kind = c.kind;                              // create only
    if (cur.match?.length) next.match = cur.match;
    if (c.verb !== "cancel") next.on = cur.on ?? { eq: "" };     // create + move keep a date
    if (cur.tolerance) next.tolerance = cur.tolerance;
    replaceAt(i, next);
  }

  function setName(i: number, name: string) { patch(i, { [opVerb(ops[i])]: name } as Partial<Op>); }

  function setMatch(i: number, raw: string) {
    const kw = raw.split(",").map((s) => s.trim()).filter(Boolean);
    patch(i, { match: kw.length ? kw : undefined });
  }

  function setPredicate(i: number, op: keyof Predicate, exprRaw: string) {
    const pred: Predicate = op === "any_of"
      ? { any_of: exprRaw.split(",").map((s) => s.trim()).filter(Boolean) }
      : { [op]: exprRaw.trim() } as Predicate;
    patch(i, { on: pred });
  }

  function addOp() { setOps([...ops, { create: "", kind: "event", on: { eq: "" } }]); }

  return (
    <div className="rounded-lg border border-slate-800 bg-slate-900/40 p-3">
      <div className="mb-2 flex items-center justify-between">
        <h3 className="text-xs font-semibold uppercase tracking-wide text-slate-400">Answer key</h3>
        <label className="flex items-center gap-1.5 text-xs text-slate-400">
          <input type="checkbox" checked={noAction} onChange={(e) => setOps(e.target.checked ? [] : [{ create: "", kind: "event", on: { eq: "" } }])} />
          this email needs no action
        </label>
      </div>

      {noAction ? (
        <p className="rounded bg-slate-800/60 px-3 py-2 text-xs text-slate-400">
          This email is graded as <strong>do nothing</strong> — any event/todo the model creates for it counts as a failure.
        </p>
      ) : (
        <div className="space-y-3">
          {ops.map((op, i) => {
            const verb = opVerb(op);
            const name = opName(op);
            const isCancel = verb === "cancel";
            const kindWord = op.kind === "todo" ? "to-do" : "event";
            const effMatch = op.match?.length ? op.match : (name ? [name] : []);
            const pred = op.on;
            return (
              <div key={i} className="rounded-md border border-slate-800 bg-slate-900 p-2.5">
                <div className="mb-2 flex items-center gap-2">
                  <select value={opChoiceId(op)} onChange={(e) => setChoice(i, e.target.value)}
                    className="rounded border border-slate-700 bg-slate-800 px-2 py-1 text-xs text-slate-200">
                    {OP_CHOICES.map((c) => <option key={c.id} value={c.id}>{c.label}</option>)}
                  </select>
                  <input value={name} onChange={(e) => setName(i, e.target.value)}
                    placeholder={`what to call this ${kindWord} (e.g. kickoff)`} className="flex-1 rounded border border-slate-700 bg-slate-800 px-2 py-1 text-xs text-slate-200" />
                  <button onClick={() => setOps(ops.filter((_, j) => j !== i))} className="px-1 text-xs text-slate-500 hover:text-rose-400">✕</button>
                </div>
                <p className="-mt-1 mb-2 text-[11px] leading-snug text-slate-500">
                  A short label for this {kindWord}. It&apos;s how a <em>later</em> email refers back to it (&ldquo;move the {name || "kickoff"}&rdquo;), and — unless you set keywords below — it&apos;s also the word we look for in the AI&apos;s title.
                </p>

                {!isCancel && (
                  <div className="mb-2 space-y-1 text-xs">
                    <div className="flex flex-wrap items-center gap-2">
                      <span className="text-slate-500">{op.kind === "todo" ? "due date" : "date"}</span>
                      <select value={predOp(pred)} onChange={(e) => setPredicate(i, e.target.value as keyof Predicate, predExpr(pred))}
                        className="rounded border border-slate-700 bg-slate-800 px-2 py-1 text-slate-200">
                        {PRED_OPS.map((o) => <option key={o.key} value={o.key}>{o.label}</option>)}
                      </select>
                      <input value={predExpr(pred)} onChange={(e) => setPredicate(i, predOp(pred), e.target.value)}
                        placeholder="e.g. next:TUE, serve+5d, or @signing+2w"
                        className="min-w-[14rem] flex-1 rounded border border-slate-700 bg-slate-800 px-2 py-1 font-mono text-sky-300 outline-none focus:border-sky-500" />
                      <button onClick={() => setPicking(i)} className="rounded bg-emerald-600/90 px-2 py-1 font-medium text-white hover:bg-emerald-500">build date</button>
                    </div>
                    <DatePreview expr={predExpr(pred)} serveDate={serveDate} />
                  </div>
                )}

                {!isCancel && /@\w/.test(predExpr(pred)) && (
                  <div className="mb-2 rounded bg-violet-500/10 px-2 py-1 text-[11px] leading-snug text-violet-300">
                    ↳ this date uses an <code className="font-mono">@anchor</code> from another email, so this email is a <strong>needle</strong> — a retrieval test. The model has to find that earlier email to answer it, and it gets harder the more filler sits between the two.
                  </div>
                )}

                {(showMatch[i] ?? !!op.match?.length) ? (
                  <div className="space-y-1 text-xs">
                    <input value={(op.match ?? []).join(", ")} onChange={(e) => setMatch(i, e.target.value)}
                      placeholder={name ? `words to find it by (optional — defaults to "${name}")` : "words to find it by (optional)"}
                      className="w-full rounded border border-slate-700 bg-slate-800 px-2 py-1 text-slate-200" />
                    <p className="text-[11px] leading-snug text-slate-500">
                      {isCancel
                        ? <>How we grade: it <strong>passes</strong> if <strong>no</strong> {kindWord} whose title contains {effMatch.length ? <code className="font-mono text-slate-400">{effMatch.join(" + ")}</code> : "this name"} is left on the calendar.</>
                        : <>How we grade: we read the {kindWord} the AI created and check its title contains {effMatch.length ? <code className="font-mono text-slate-400">{effMatch.join(" + ")}</code> : "this name"} (any capitalization), on the right day. Only fill this in if the AI&apos;s natural title wouldn&apos;t contain the name (e.g. name <code className="font-mono text-slate-400">filing</code> → keyword <code className="font-mono text-slate-400">HSR</code>).</>}
                    </p>
                  </div>
                ) : (
                  <button onClick={() => setShowMatch((m) => ({ ...m, [i]: true }))} className="text-[11px] text-slate-500 hover:text-sky-300">▸ the AI might title it differently</button>
                )}

                {!isCancel && (
                  <label className="mt-2 flex items-center gap-1.5 text-xs text-slate-400">how close does the day have to be?
                    <select value={op.tolerance ?? "exact_day"} onChange={(e) => patch(i, { tolerance: e.target.value === "exact_day" ? undefined : e.target.value })}
                      className="rounded border border-slate-700 bg-slate-800 px-2 py-1 text-slate-200">
                      {!TOLERANCE_CHOICES.some((c) => c.value === (op.tolerance ?? "exact_day")) &&
                        <option value={op.tolerance}>{op.tolerance}</option>}
                      {TOLERANCE_CHOICES.map((c) => <option key={c.value} value={c.value}>{c.label}</option>)}
                    </select>
                  </label>
                )}

                {picking === i && !isCancel && (
                  <TokenBlockly anchors={anchors} serveDate={serveDate} mode="predicate"
                    onInsert={(expr) => { setPredicate(i, predOp(pred), predOp(pred) === "any_of" ? `${predExpr(pred)}${predExpr(pred) ? ", " : ""}${expr}` : expr); setPicking(null); }}
                    onClose={() => setPicking(null)} />
                )}
              </div>
            );
          })}
          <button onClick={addOp} className="rounded-md border border-dashed border-slate-700 px-3 py-1.5 text-xs text-slate-400 hover:border-sky-600 hover:text-sky-300">+ another action</button>
        </div>
      )}
    </div>
  );
}
