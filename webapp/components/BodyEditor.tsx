"use client";
import { useEffect, useRef, useState } from "react";
import { resolveExpr } from "@/lib/api";
import { splitBody } from "@/lib/grammar";
import { parseExpr } from "@/lib/dateExpr";
import DateBuilder from "./DateBuilder";

interface Props {
  body: string;
  anchors: string[];
  serveDate: string;
  onChange: (body: string) => void;
}

// pull the resolvable expression out of a {token} or {!name = expr} span
function innerExpr(token: string): { expr: string; emit?: string } {
  const inner = token.replace(/^\{|\}$/g, "").trim();
  const m = inner.match(/^!\s*([A-Za-z_]\w*)\s*=\s*(.+)$/);
  return m ? { expr: m[2].trim(), emit: m[1] } : { expr: inner };
}

export default function BodyEditor({ body, anchors, serveDate, onChange }: Props) {
  const [picking, setPicking] = useState(false);
  const [draft, setDraft] = useState("");                 // the expression being built for insertion
  const [emit, setEmit] = useState(false);                // publish this date as a named anchor?
  const [emitName, setEmitName] = useState("");
  const taRef = useRef<HTMLTextAreaElement>(null);
  const caret = useRef<number | null>(null);              // cursor position captured when the panel opens
  const [chips, setChips] = useState<Record<string, string>>({});

  // resolve each distinct token for the inline chips
  useEffect(() => {
    const tokens = splitBody(body).filter((s) => s.token).map((s) => s.text);
    const uniq = [...new Set(tokens)];
    let alive = true;
    (async () => {
      const out: Record<string, string> = {};
      for (const tok of uniq) {
        const { expr, emit } = innerExpr(tok);
        if (/@\w/.test(expr)) { out[tok] = "resolves at serve time"; continue; }
        const r = await resolveExpr(expr, serveDate, {});
        out[tok] = r.ok ? (emit ? `${emit} = ${r.human}` : r.human ?? "") : `⚠ ${r.error}`;
      }
      if (alive) setChips(out);
    })();
    return () => { alive = false; };
  }, [body, serveDate]);

  function openPanel() {
    const ta = taRef.current;
    caret.current = ta ? ta.selectionStart : body.length;   // remember where to drop the token
    setDraft(""); setEmit(false); setEmitName(""); setPicking(true);
  }
  function insert() {
    const token = emit && emitName ? `{!${emitName} = ${draft}}` : `{${draft}}`;
    const at = caret.current ?? body.length;
    onChange(body.slice(0, at) + token + body.slice(at));
    setPicking(false);
  }

  const nameOk = !emit || /^[A-Za-z_]\w*$/.test(emitName);
  const canInsert = !!draft && parseExpr(draft) !== null && nameOk;
  const spans = splitBody(body);

  return (
    <div>
      <div className="mb-1 flex items-center justify-between">
        <label className="text-xs font-semibold uppercase tracking-wide text-slate-400">Body</label>
        <button onClick={() => (picking ? setPicking(false) : openPanel())}
          className="rounded-md bg-emerald-600/90 px-2.5 py-1 text-xs font-medium text-white hover:bg-emerald-500">+ insert date</button>
      </div>
      <textarea ref={taRef} value={body} onChange={(e) => onChange(e.target.value)} rows={6}
        placeholder="Write the email. Use “+ insert date” for any date, so the email and the answer key always resolve to the same day."
        className="w-full resize-y rounded-md border border-slate-700 bg-slate-800 px-3 py-2 font-mono text-sm text-slate-100 outline-none focus:border-sky-500" />

      {picking && (
        <div className="mt-2 space-y-2 rounded-lg border border-slate-700 bg-slate-900 p-3">
          <div className="flex items-center justify-between">
            <span className="text-xs font-semibold text-slate-200">Insert a date</span>
            <button onClick={() => setPicking(false)} className="text-xs text-slate-500 hover:text-slate-300">cancel</button>
          </div>
          <DateBuilder value={draft} anchors={anchors} serveDate={serveDate} onChange={setDraft} />
          <label className="flex flex-wrap items-center gap-2 text-xs text-slate-300">
            <input type="checkbox" checked={emit} onChange={(e) => setEmit(e.target.checked)} />
            other emails can refer to this date as
            <input value={emitName} onChange={(e) => setEmitName(e.target.value)} disabled={!emit}
              placeholder="signing" className="w-28 rounded border border-slate-700 bg-slate-800 px-2 py-0.5 font-mono text-xs disabled:opacity-40" />
          </label>
          {emit && emitName && !nameOk && <p className="text-[11px] text-rose-400">use letters, numbers and underscores (start with a letter or _)</p>}
          <div className="flex justify-end">
            <button onClick={insert} disabled={!canInsert}
              className="rounded-md bg-sky-600 px-4 py-1.5 text-sm font-medium text-white hover:bg-sky-500 disabled:opacity-40">Insert</button>
          </div>
        </div>
      )}

      {spans.some((s) => s.token) && (
        <div className="mt-2 flex flex-wrap gap-1.5">
          {[...new Set(spans.filter((s) => s.token).map((s) => s.text))].map((tok) => (
            <span key={tok} className="inline-flex items-center gap-1 rounded-md border border-slate-700 bg-slate-900 px-2 py-0.5 text-xs">
              <code className="font-mono text-sky-300">{tok}</code>
              <span className={chips[tok]?.startsWith("⚠") ? "text-rose-400" : "text-slate-400"}>→ {chips[tok] ?? "…"}</span>
            </span>
          ))}
        </div>
      )}
    </div>
  );
}
