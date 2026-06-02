"use client";
import { useState } from "react";
import { OP_CHOICES, opChoiceId, opName, opVerb } from "@/lib/grammar";
import type { Answer, Op, Predicate } from "@/lib/types";
import TokenBlockly from "./blockly/TokenBlockly";

interface Props {
  answer: Answer;
  anchors: string[];
  serveDate: string;
  onChange: (answer: Answer) => void;
}

const PRED_OPS: { key: keyof Predicate; label: string }[] = [
  { key: "eq", label: "on exactly (eq)" },
  { key: "by", label: "on or before (by)" },
  { key: "in", label: "within interval (in)" },
  { key: "any_of", label: "any of (any_of)" },
  { key: "not_in", label: "not within (not_in)" },
];

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
          no action expected (FYI / distractor)
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
                    placeholder="obligation name (e.g. kickoff)" className="flex-1 rounded border border-slate-700 bg-slate-800 px-2 py-1 text-xs text-slate-200" />
                  <button onClick={() => setOps(ops.filter((_, j) => j !== i))} className="px-1 text-xs text-slate-500 hover:text-rose-400">✕</button>
                </div>

                {!isCancel && (
                  <div className="mb-2 flex flex-wrap items-center gap-2 text-xs">
                    <span className="text-slate-500">{op.kind === "todo" ? "due date" : "date"}</span>
                    <select value={predOp(pred)} onChange={(e) => setPredicate(i, e.target.value as keyof Predicate, predExpr(pred))}
                      className="rounded border border-slate-700 bg-slate-800 px-2 py-1 text-slate-200">
                      {PRED_OPS.map((o) => <option key={o.key} value={o.key}>{o.label}</option>)}
                    </select>
                    <code className="rounded bg-slate-800 px-2 py-1 font-mono text-sky-300">{predExpr(pred) || "— no date —"}</code>
                    <button onClick={() => setPicking(i)} className="rounded bg-emerald-600/90 px-2 py-1 font-medium text-white hover:bg-emerald-500">build date</button>
                  </div>
                )}

                {!isCancel && /@\w/.test(predExpr(pred)) && (
                  <div className="mb-2 rounded bg-violet-500/10 px-2 py-1 text-[11px] leading-snug text-violet-300">
                    ↳ this date uses an <code className="font-mono">@anchor</code> from another email, so this email is a <strong>needle</strong> — a retrieval test. The model has to find that earlier email to answer it, and it gets harder the more filler sits between the two.
                  </div>
                )}

                <div className="space-y-1 text-xs">
                  <input value={(op.match ?? []).join(", ")} onChange={(e) => setMatch(i, e.target.value)}
                    placeholder={name ? `title keywords (default: ${name})` : "title keywords"}
                    className="w-full rounded border border-slate-700 bg-slate-800 px-2 py-1 text-slate-200" />
                  <p className="text-[11px] leading-snug text-slate-500">
                    {isCancel
                      ? <>The grader checks that no {kindWord} matching {effMatch.length ? <code className="font-mono text-slate-400">{effMatch.join(" + ")}</code> : "this obligation"} is left on the calendar.</>
                      : <>We find the model&apos;s {kindWord} by matching its title against {effMatch.length ? <code className="font-mono text-slate-400">{effMatch.join(" + ")}</code> : "the obligation name"}. Leave blank to use the name. Pick a word a natural title would actually contain.</>}
                  </p>
                </div>

                {!isCancel && (
                  <label className="mt-2 flex items-center gap-1 text-xs text-slate-400">tolerance
                    <input value={op.tolerance ?? "exact_day"} onChange={(e) => patch(i, { tolerance: e.target.value })}
                      className="w-28 rounded border border-slate-700 bg-slate-800 px-2 py-1 text-slate-200" />
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
          <button onClick={addOp} className="rounded-md border border-dashed border-slate-700 px-3 py-1.5 text-xs text-slate-400 hover:border-sky-600 hover:text-sky-300">+ obligation</button>
        </div>
      )}
    </div>
  );
}
