"use client";
import { useEffect, useRef, useState } from "react";
import { resolveExpr } from "@/lib/api";
import type { ResolveResult } from "@/lib/types";
import { compileWorkspace, defineGrammarBlocks, setAnchors, TOOLBOX } from "./blocks";

interface Props {
  anchors: string[];
  serveDate: string;
  // "token" wraps the result for an email body (optionally as an emission);
  // "predicate" returns the bare expression for an answer-key date slot.
  mode: "token" | "predicate";
  onInsert: (value: string) => void;
  onClose: () => void;
}

export default function TokenBlockly({ anchors, serveDate, mode, onInsert, onClose }: Props) {
  const host = useRef<HTMLDivElement>(null);
  const wsRef = useRef<any>(null);
  const [expr, setExpr] = useState<string | null>(null);
  const [preview, setPreview] = useState<ResolveResult | null>(null);
  const [emit, setEmit] = useState(false);
  const [emitName, setEmitName] = useState("");

  useEffect(() => {
    let ws: any;
    let cancelled = false;
    (async () => {
      const Blockly = await import("blockly");
      if (cancelled || !host.current) return;
      defineGrammarBlocks(Blockly);
      setAnchors(anchors);
      ws = Blockly.inject(host.current, {
        toolbox: TOOLBOX,
        renderer: "zelos", // the Scratch-style look
        theme: (Blockly as any).Themes?.Zelos ?? undefined,
        trashcan: true,
        scrollbars: true,
        zoom: { controls: true, wheel: true, startScale: 0.95 },
      });
      wsRef.current = ws;
      ws.addChangeListener(() => {
        const e = compileWorkspace(ws);
        setExpr(e);
      });
    })();
    return () => {
      cancelled = true;
      try { ws?.dispose(); } catch { /* noop */ }
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Live-resolve the built expression against the real Python resolver.
  useEffect(() => {
    if (!expr) { setPreview(null); return; }
    let alive = true;
    resolveExpr(expr, serveDate, {}).then((r) => alive && setPreview(r));
    return () => { alive = false; };
  }, [expr, serveDate]);

  const anchorRefShown = expr && /@\w/.test(expr) && preview && !preview.ok;
  const canInsert = !!expr && (mode !== "token" || !emit || /^[A-Za-z_]\w*$/.test(emitName));

  function doInsert() {
    if (!expr) return;
    if (mode === "predicate") return onInsert(expr);
    onInsert(emit && emitName ? `{!${emitName} = ${expr}}` : `{${expr}}`);
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4">
      <div className="flex h-[80vh] w-[min(1100px,95vw)] flex-col overflow-hidden rounded-xl border border-slate-700 bg-slate-900 shadow-2xl">
        <div className="flex items-center justify-between border-b border-slate-800 px-4 py-3">
          <div>
            <h2 className="text-sm font-semibold text-slate-100">Date token builder</h2>
            <p className="text-xs text-slate-400">Snap blocks together — the date resolves with the real grader so what you see is what gets graded.</p>
          </div>
          <button onClick={onClose} className="rounded-md px-2 py-1 text-sm text-slate-400 hover:bg-slate-800">✕</button>
        </div>

        <div className="relative min-h-0 flex-1">
          <div ref={host} className="h-full w-full" />
          {!expr && (
            <div className="pointer-events-none absolute inset-0 flex items-center justify-center">
              <div className="rounded-lg border border-dashed border-slate-600 bg-slate-900/70 px-4 py-3 text-center text-sm text-slate-400">
                Drag a block from the left to start.<br />
                <span className="text-xs text-slate-500">Most dates begin with <span className="font-mono text-sky-300">next ·</span>, <span className="font-mono text-sky-300">this ·</span> or <span className="font-mono text-sky-300">serve</span>.</span>
              </div>
            </div>
          )}
        </div>

        <div className="border-t border-slate-800 px-4 py-3">
          <div className="mb-2 flex flex-wrap items-center gap-3 text-sm">
            <span className="text-slate-400">Expression:</span>
            <code className="rounded bg-slate-800 px-2 py-0.5 font-mono text-sky-300">{expr ?? "build a block stack…"}</code>
            {preview?.ok && <span className="text-emerald-400">→ {preview.human}</span>}
            {expr && preview && !preview.ok && !anchorRefShown && <span className="text-rose-400">→ {preview.error}</span>}
            {anchorRefShown && <span className="text-amber-400">→ uses an anchor; resolves at serve time</span>}
          </div>

          <div className="flex items-center justify-between gap-4">
            {mode === "token" ? (
              <label className="flex items-center gap-2 text-sm text-slate-300">
                <input type="checkbox" checked={emit} onChange={(e) => setEmit(e.target.checked)} />
                publish this date as an anchor named
                <input value={emitName} onChange={(e) => setEmitName(e.target.value)} disabled={!emit}
                  placeholder="signing" className="w-28 rounded border border-slate-700 bg-slate-800 px-2 py-0.5 font-mono text-xs disabled:opacity-40" />
              </label>
            ) : <span className="text-xs text-slate-500">This expression becomes the graded date for the answer-key slot.</span>}
            <button onClick={doInsert} disabled={!canInsert}
              className="rounded-md bg-sky-600 px-4 py-1.5 text-sm font-medium text-white hover:bg-sky-500 disabled:opacity-40">
              Insert
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
