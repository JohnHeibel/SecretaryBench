"use client";
import { useState } from "react";
import { OP_CHOICES, opChoiceId, opName, opVerb } from "@/lib/grammar";
import type { Answer, Op, Predicate } from "@/lib/types";
import DateBuilder from "./DateBuilder";

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
  { key: "eq", label: "on exactly" },
  { key: "by", label: "on or before" },
  { key: "in", label: "within" },
  { key: "any_of", label: "any of" },
  { key: "not_in", label: "not within" },
];

function predOp(p?: Predicate): keyof Predicate {
  if (!p) return "eq";
  return (Object.keys(p)[0] as keyof Predicate) ?? "eq";
}
// The single date expression for a non-any_of predicate (or the first option of an any_of).
function firstExpr(p?: Predicate): string {
  if (!p) return "";
  const v = Object.values(p)[0];
  return Array.isArray(v) ? (v[0] ?? "") : (v ?? "");
}
// The list of expressions for an any_of predicate (folding a single-expr predicate to a 1-list).
function anyOfList(p?: Predicate): string[] {
  if (!p) return [];
  return Array.isArray(p.any_of) ? p.any_of : (firstExpr(p) ? [firstExpr(p)] : []);
}

export default function AnswerKeyBuilder({ answer, anchors, serveDate, onChange }: Props) {
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
                  <div className="mb-2 space-y-1.5 text-xs">
                    <div className="flex flex-wrap items-center gap-2">
                      <span className="text-slate-500">{op.kind === "todo" ? "due date" : "date"}</span>
                      <select value={predOp(pred)} onChange={(e) => {
                        const newOp = e.target.value as keyof Predicate;
                        if (newOp === "any_of") patch(i, { on: { any_of: anyOfList(pred) } });
                        else setPredicate(i, newOp, firstExpr(pred));
                      }} className="rounded border border-slate-700 bg-slate-800 px-2 py-1 text-slate-200">
                        {PRED_OPS.map((o) => <option key={o.key} value={o.key}>{o.label}</option>)}
                      </select>
                    </div>
                    {predOp(pred) === "any_of" ? (
                      <AnyOfBuilder values={anyOfList(pred)} anchors={anchors} serveDate={serveDate}
                        onChange={(list) => patch(i, { on: { any_of: list } })} />
                    ) : (
                      <DateBuilder value={firstExpr(pred)} anchors={anchors} serveDate={serveDate}
                        intervalOnly={predOp(pred) === "in" || predOp(pred) === "not_in"}
                        onChange={(expr) => setPredicate(i, predOp(pred), expr)} />
                    )}
                  </div>
                )}

                {!isCancel && anyOfList(pred).some((e) => /@\w/.test(e)) && (
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

              </div>
            );
          })}
          <button onClick={addOp} className="rounded-md border border-dashed border-slate-700 px-3 py-1.5 text-xs text-slate-400 hover:border-sky-600 hover:text-sky-300">+ another action</button>
        </div>
      )}
    </div>
  );
}

// "any of" = a list of acceptable dates (the model may land on any one). Each is its own
// DateBuilder; an empty list shows a single blank builder to start from.
function AnyOfBuilder({ values, anchors, serveDate, onChange }: { values: string[]; anchors: string[]; serveDate: string; onChange: (list: string[]) => void }) {
  const list = values.length ? values : [""];
  return (
    <div className="space-y-1.5">
      {list.map((v, k) => (
        <div key={k} className="flex items-start gap-2">
          <div className="flex-1"><DateBuilder value={v} anchors={anchors} serveDate={serveDate} onChange={(e) => onChange(list.map((x, j) => j === k ? e : x))} /></div>
          {list.length > 1 && <button type="button" onClick={() => onChange(list.filter((_, j) => j !== k))} className="mt-1 px-1 text-slate-500 hover:text-rose-400">✕</button>}
        </div>
      ))}
      <button type="button" onClick={() => onChange([...list, ""])} className="text-[11px] text-slate-500 hover:text-sky-300">+ add option</button>
    </div>
  );
}
