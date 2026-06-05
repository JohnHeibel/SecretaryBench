"use client";
import { useState } from "react";
import { OP_CHOICES, opChoiceId, opName, opVerb } from "@/lib/grammar";
import { parseExpr, serializeExpr } from "@/lib/dateExpr";
import type { Answer, ObjKind, Op, Predicate } from "@/lib/types";
import DateBuilder from "./DateBuilder";

interface Props {
  answer: Answer;
  anchors: string[];
  serveDate: string;
  obligations: string[];                          // obligation names created across this node — what a move/cancel targets, and reuses as @anchors
  obligationKinds: Record<string, ObjKind>;       // name -> kind, so a move/cancel (which stores no kind) knows it targets an event vs to-do
  onChange: (answer: Answer) => void;
}

const ANCHOR_ID = /^[A-Za-z_][A-Za-z0-9_]*$/;   // a name usable as @NAME in the grammar (no spaces)

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
// Every date expression in a predicate (any key) — used to spot an @anchor reference (a needle).
function predExprs(p?: Predicate): string[] {
  if (!p) return [];
  return Object.values(p).flatMap((v) => Array.isArray(v) ? v : [v]).filter(Boolean) as string[];
}
const usesAnchor = (p?: Predicate) => predExprs(p).some((e) => /@[A-Za-z_]/.test(e));
// Drop any clock suffix from an expr string — used when a date moves to a slot that grades day-level
// (a `by` deadline, an `in`/`not_in` window), so a stale, now-uneditable time can't linger.
const dropTime = (expr: string) => { const p = parseExpr(expr); return p?.time ? serializeExpr({ ...p, time: undefined }) : expr; };

export default function AnswerKeyBuilder({ answer, anchors, serveDate, obligations, obligationKinds, onChange }: Props) {
  // which op indices have the (advanced) keyword override revealed
  const [showMatch, setShowMatch] = useState<Record<number, boolean>>({});
  const ops = answer.ops ?? [];
  const noAction = ops.length === 0;

  function setOps(next: Op[]) { onChange({ ...answer, ops: next }); }
  function patch(i: number, p: Partial<Op>) { setOps(ops.map((o, j) => j === i ? { ...o, ...p } : o)); }
  function replaceAt(i: number, op: Op) { setOps(ops.map((o, j) => j === i ? op : o)); }

  // Changing the verb means the JSON KEY changes (create -> move), so rebuild the op from
  // scratch, carrying the name/match/date and dropping whatever the new verb can't hold. Tolerance
  // is never carried — we grade exact day/time now (the within:Nd knob is retired).
  function setChoice(i: number, choiceId: string) {
    const c = OP_CHOICES.find((x) => x.id === choiceId);
    if (!c) return;
    const cur = ops[i];
    const next: Op = { [c.verb]: opName(cur) };
    if (c.kind) next.kind = c.kind;                              // create only
    if (cur.match?.length) next.match = cur.match;
    if (c.verb !== "cancel") next.on = cur.on ?? { eq: "" };     // create + move keep a date
    replaceAt(i, next);
  }

  // Obligations from OTHER ops are reusable as @anchors for this op's date (a later email reuses an
  // earlier date — the long-horizon needle); exclude this op's own name so a create can't cite itself.
  // Only identifier-safe names can be an @anchor (the grammar's @NAME has no spaces), so filter the rest.
  const anchorsFor = (self: string) => [...new Set([...obligations.filter((o) => o && o !== self && ANCHOR_ID.test(o)), ...anchors])];

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
      <h3 className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-400">Answer key</h3>

      {noAction ? (
        <p className="rounded bg-slate-800/60 px-3 py-2 text-xs text-slate-400">
          This is an <strong>FYI / Junk</strong> email — graded as <strong>do nothing</strong>. Any event or to-do the model creates for it counts as a failure. To give it an action instead, set its type to <strong>Action</strong> at the top.
        </p>
      ) : (
        <div className="space-y-3">
          {ops.map((op, i) => {
            const verb = opVerb(op);
            const name = opName(op);
            const isCreate = verb === "create";
            const isCancel = verb === "cancel";
            const ek = op.kind ?? obligationKinds[name];   // effective kind (a move/cancel inherits its create's kind)
            const kindWord = ek === "todo" ? "to-do" : "event";
            const effMatch = op.match?.length ? op.match : (name ? [name] : []);
            const pred = op.on;
            const po = predOp(pred);
            return (
              <div key={i} className="rounded-md border border-slate-800 bg-slate-900 p-2.5">
                <div className="mb-2 flex items-center gap-2">
                  <select value={opChoiceId(op)} onChange={(e) => setChoice(i, e.target.value)}
                    className="rounded border border-slate-700 bg-slate-800 px-2 py-1 text-xs text-slate-200">
                    {OP_CHOICES.map((c) => <option key={c.id} value={c.id}>{c.label}</option>)}
                  </select>
                  {isCreate ? (
                    <input value={name} onChange={(e) => setName(i, e.target.value)}
                      placeholder={`What's this ${kindWord} called? (e.g. kickoff)`} className="flex-1 rounded border border-slate-700 bg-slate-800 px-2 py-1 text-xs text-slate-200" />
                  ) : obligations.length ? (
                    <select value={name} onChange={(e) => setName(i, e.target.value)} className="flex-1 rounded border border-slate-700 bg-slate-800 px-2 py-1 text-xs text-slate-200">
                      <option value="" disabled>which one?</option>
                      {obligations.map((o) => <option key={o} value={o}>{o}</option>)}
                      {name && !obligations.includes(name) && <option value={name}>{name} (not created yet)</option>}
                    </select>
                  ) : (
                    <input value={name} onChange={(e) => setName(i, e.target.value)}
                      placeholder={`name of the thing to ${verb}`} className="flex-1 rounded border border-slate-700 bg-slate-800 px-2 py-1 text-xs text-slate-200" />
                  )}
                  <button onClick={() => setOps(ops.filter((_, j) => j !== i))} className="px-1 text-xs text-slate-500 hover:text-rose-400">✕</button>
                </div>
                <p className="-mt-1 mb-2 text-[11px] leading-snug text-slate-500">
                  {isCreate
                    ? <>The name the email uses for it. A <em>later</em> email can move or cancel it by this name (&ldquo;move the {name || "kickoff"}&rdquo;) and reuse its date.</>
                    : <>{verb === "move" ? "Reschedules" : "Cancels"} the {kindWord} an earlier email created — pick it by name above.</>}
                </p>

                {!isCancel && (
                  <div className="mb-2 space-y-1.5 text-xs">
                    <div className="flex flex-wrap items-center gap-2">
                      <span className="text-slate-500">{ek === "todo" ? "due date" : "date"}</span>
                      <select value={po} onChange={(e) => {
                        const newOp = e.target.value as keyof Predicate;
                        if (newOp === "any_of") patch(i, { on: { any_of: anyOfList(pred) } });
                        else setPredicate(i, newOp, newOp === "eq" ? firstExpr(pred) : dropTime(firstExpr(pred)));
                      }} className="rounded border border-slate-700 bg-slate-800 px-2 py-1 text-slate-200">
                        {PRED_OPS.map((o) => <option key={o.key} value={o.key}>{o.label}</option>)}
                      </select>
                    </div>
                    {po === "any_of" ? (
                      <AnyOfBuilder values={anyOfList(pred)} anchors={anchorsFor(name)} serveDate={serveDate} allowTime={ek === "event"}
                        onChange={(list) => patch(i, { on: { any_of: list } })} />
                    ) : (
                      <DateBuilder value={firstExpr(pred)} anchors={anchorsFor(name)} serveDate={serveDate}
                        intervalOnly={po === "in" || po === "not_in"} allowTime={ek === "event" && po === "eq"}
                        onChange={(expr) => setPredicate(i, po, expr)} />
                    )}
                  </div>
                )}

                {!isCancel && usesAnchor(pred) && (
                  <div className="mb-2 rounded bg-violet-500/10 px-2 py-1 text-[11px] leading-snug text-violet-300">
                    ↳ this date reuses an <code className="font-mono">@anchor</code> from an earlier email, so this is a <strong>needle</strong> — a retrieval test. The model has to recall that earlier email to answer it, and it gets harder the more filler sits between the two. (The dependency link is added for you.)
                  </div>
                )}

                {isCreate && ((showMatch[i] ?? !!op.match?.length) ? (
                  <div className="space-y-1 text-xs">
                    <input value={(op.match ?? []).join(", ")} onChange={(e) => setMatch(i, e.target.value)}
                      placeholder={name ? `title keywords (optional — defaults to "${name}")` : "title keywords (optional)"}
                      className="w-full rounded border border-slate-700 bg-slate-800 px-2 py-1 text-slate-200" />
                    <p className="text-[11px] leading-snug text-slate-500">
                      Advanced: we find the AI&apos;s {kindWord} by checking its title contains {effMatch.length ? <code className="font-mono text-slate-400">{effMatch.join(" + ")}</code> : "this name"} (any capitalization). Only set this if the AI&apos;s natural title wouldn&apos;t contain the name (e.g. name <code className="font-mono text-slate-400">filing</code> → keyword <code className="font-mono text-slate-400">HSR</code>).
                    </p>
                  </div>
                ) : (
                  <button onClick={() => setShowMatch((m) => ({ ...m, [i]: true }))} className="text-[11px] text-slate-500 hover:text-sky-300">▸ Advanced</button>
                ))}

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
function AnyOfBuilder({ values, anchors, serveDate, allowTime, onChange }: { values: string[]; anchors: string[]; serveDate: string; allowTime?: boolean; onChange: (list: string[]) => void }) {
  const list = values.length ? values : [""];
  return (
    <div className="space-y-1.5">
      {list.map((v, k) => (
        <div key={k} className="flex items-start gap-2">
          <div className="flex-1"><DateBuilder value={v} anchors={anchors} serveDate={serveDate} allowTime={allowTime} onChange={(e) => onChange(list.map((x, j) => j === k ? e : x))} /></div>
          {list.length > 1 && <button type="button" onClick={() => onChange(list.filter((_, j) => j !== k))} className="mt-1 px-1 text-slate-500 hover:text-rose-400">✕</button>}
        </div>
      ))}
      <button type="button" onClick={() => onChange([...list, ""])} className="text-[11px] text-slate-500 hover:text-sky-300">+ add option</button>
    </div>
  );
}
