"use client";
import { useState } from "react";
import { ACTIONS } from "@/lib/grammar";
import type { Answer, ExpectEntry, Predicate } from "@/lib/types";
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

function dateField(action: string): "start" | "due" | null {
  if (action === "create_todo") return "due";
  if (action === "reply" || action === "delegate") return null;
  return "start";
}

export default function AnswerKeyBuilder({ answer, anchors, serveDate, onChange }: Props) {
  // which expect-entry index is currently picking a date via the block builder
  const [picking, setPicking] = useState<number | null>(null);
  const noAction = answer.expect.length === 0 && (answer.forbid?.length ?? 0) > 0;

  function setExpect(next: ExpectEntry[]) { onChange({ ...answer, expect: next }); }
  function patch(i: number, p: Partial<ExpectEntry>) { setExpect(answer.expect.map((e, j) => j === i ? { ...e, ...p } : e)); }

  function setPredicate(i: number, op: keyof Predicate, exprRaw: string) {
    const f = dateField(answer.expect[i].action);
    if (!f) return;
    const pred: Predicate = op === "any_of"
      ? { any_of: exprRaw.split(",").map((s) => s.trim()).filter(Boolean) }
      : { [op]: exprRaw.trim() } as Predicate;
    patch(i, { [f]: pred } as Partial<ExpectEntry>);
  }

  function addEntry() {
    setExpect([...answer.expect, { action: "create_event", title_match: [], start: { eq: "" }, count: 1, tolerance: "exact_day" }]);
  }

  function toggleNoAction(on: boolean) {
    onChange(on ? { expect: [], forbid: [{}] } : { expect: [], forbid: [] });
  }

  return (
    <div className="rounded-lg border border-slate-800 bg-slate-900/40 p-3">
      <div className="mb-2 flex items-center justify-between">
        <h3 className="text-xs font-semibold uppercase tracking-wide text-slate-400">Answer key</h3>
        <label className="flex items-center gap-1.5 text-xs text-slate-400">
          <input type="checkbox" checked={noAction} onChange={(e) => toggleNoAction(e.target.checked)} />
          no action expected (FYI / distractor)
        </label>
      </div>

      {noAction ? (
        <p className="rounded bg-slate-800/60 px-3 py-2 text-xs text-slate-400">
          This email is graded as <strong>do nothing</strong> — any event/todo the model creates for it counts as a failure.
        </p>
      ) : (
        <div className="space-y-3">
          {answer.expect.map((entry, i) => {
            const f = dateField(entry.action);
            const pred = f ? (entry[f] as Predicate | undefined) : undefined;
            return (
              <div key={i} className="rounded-md border border-slate-800 bg-slate-900 p-2.5">
                <div className="mb-2 flex items-center gap-2">
                  <select value={entry.action} onChange={(e) => patch(i, { action: e.target.value as ExpectEntry["action"] })}
                    className="rounded border border-slate-700 bg-slate-800 px-2 py-1 text-xs text-slate-200">
                    {ACTIONS.map((a) => <option key={a} value={a}>{a}</option>)}
                  </select>
                  <input value={(entry.title_match ?? []).join(", ")} onChange={(e) => patch(i, { title_match: e.target.value.split(",").map((s) => s.trim()).filter(Boolean) })}
                    placeholder="title keywords (e.g. kickoff)" className="flex-1 rounded border border-slate-700 bg-slate-800 px-2 py-1 text-xs text-slate-200" />
                  <button onClick={() => setExpect(answer.expect.filter((_, j) => j !== i))} className="px-1 text-xs text-slate-500 hover:text-rose-400">✕</button>
                </div>

                {f && (
                  <div className="mb-2 flex flex-wrap items-center gap-2 text-xs">
                    <span className="text-slate-500">{f === "due" ? "due date" : "date"}</span>
                    <select value={predOp(pred)} onChange={(e) => setPredicate(i, e.target.value as keyof Predicate, predExpr(pred))}
                      className="rounded border border-slate-700 bg-slate-800 px-2 py-1 text-slate-200">
                      {PRED_OPS.map((o) => <option key={o.key} value={o.key}>{o.label}</option>)}
                    </select>
                    <code className="rounded bg-slate-800 px-2 py-1 font-mono text-sky-300">{predExpr(pred) || "— no date —"}</code>
                    <button onClick={() => setPicking(i)} className="rounded bg-emerald-600/90 px-2 py-1 font-medium text-white hover:bg-emerald-500">build date</button>
                  </div>
                )}

                <div className="flex items-center gap-3 text-xs text-slate-400">
                  <label className="flex items-center gap-1">count
                    <input type="number" value={entry.count ?? 1} min={0} onChange={(e) => patch(i, { count: Number(e.target.value) })}
                      className="w-16 rounded border border-slate-700 bg-slate-800 px-2 py-1 text-slate-200" />
                  </label>
                  <label className="flex items-center gap-1">tolerance
                    <input value={entry.tolerance ?? "exact_day"} onChange={(e) => patch(i, { tolerance: e.target.value })}
                      className="w-28 rounded border border-slate-700 bg-slate-800 px-2 py-1 text-slate-200" />
                  </label>
                </div>

                {picking === i && f && (
                  <TokenBlockly anchors={anchors} serveDate={serveDate} mode="predicate"
                    onInsert={(expr) => { setPredicate(i, predOp(pred), predOp(pred) === "any_of" ? `${predExpr(pred)}${predExpr(pred) ? ", " : ""}${expr}` : expr); setPicking(null); }}
                    onClose={() => setPicking(null)} />
                )}
              </div>
            );
          })}
          <button onClick={addEntry} className="rounded-md border border-dashed border-slate-700 px-3 py-1.5 text-xs text-slate-400 hover:border-sky-600 hover:text-sky-300">+ expected action</button>
        </div>
      )}
    </div>
  );
}
