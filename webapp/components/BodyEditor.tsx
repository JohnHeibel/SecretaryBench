"use client";
import { useEffect, useRef, useState } from "react";
import { resolveExpr } from "@/lib/api";
import { splitBody } from "@/lib/grammar";
import TokenBlockly from "./blockly/TokenBlockly";

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
  const taRef = useRef<HTMLTextAreaElement>(null);
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

  function insertToken(value: string) {
    const ta = taRef.current;
    const at = ta ? ta.selectionStart : body.length;
    onChange(body.slice(0, at) + value + body.slice(ta ? ta.selectionEnd : body.length));
    setPicking(false);
  }

  const spans = splitBody(body);

  return (
    <div>
      <div className="mb-1 flex items-center justify-between">
        <label className="text-xs font-semibold uppercase tracking-wide text-slate-400">Body</label>
        <button onClick={() => setPicking(true)} className="rounded-md bg-emerald-600/90 px-2.5 py-1 text-xs font-medium text-white hover:bg-emerald-500">+ insert date token</button>
      </div>
      <textarea ref={taRef} value={body} onChange={(e) => onChange(e.target.value)} rows={6}
        placeholder="Write the email. Insert date tokens with the block builder so the date and the answer key can never drift."
        className="w-full resize-y rounded-md border border-slate-700 bg-slate-800 px-3 py-2 font-mono text-sm text-slate-100 outline-none focus:border-sky-500" />

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

      {picking && <TokenBlockly anchors={anchors} serveDate={serveDate} mode="token" onInsert={insertToken} onClose={() => setPicking(false)} />}
    </div>
  );
}
