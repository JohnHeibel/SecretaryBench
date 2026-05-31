import { useEffect, useRef, useState } from "react";
import { api } from "../api";
import { emptyChip } from "../chipLabel";
import type { Chip, ChipBase, EmailForm, Segment } from "../types";
import { ChipPill } from "./ChipPill";
import { FactPill, type FactSeg } from "./FactPill";

type InsertOption = {
  label: string;
  hint: string;
  make: (reachableFacts: string[]) => Segment;
  drag?: ChipBase["kind"];
};

const INSERTS: InsertOption[] = [
  { label: "📅 A date", hint: "next Thursday, in 5 days…", make: () => ({ type: "chip", chip: emptyChip("next_weekday") }), drag: "next_weekday" },
  { label: "📌 A saved date (pin)", hint: "reuse a date from earlier", make: () => ({ type: "chip", chip: emptyChip("anchor") }), drag: "anchor" },
  { label: "📏 A shared value", hint: "define a reusable value (e.g. a duration)", make: () => ({ type: "fact", name: "", value: "90m" }) },
  { label: "↩ Reuse a shared value", hint: "insert a value defined earlier", make: (rf) => ({ type: "fact", name: rf[0] || "", value: null }) },
  { label: "⚙ Advanced date", hint: "type a raw token", make: () => ({ type: "chip", chip: emptyChip("raw") }), drag: "raw" },
];

interface Props {
  email: EmailForm;
  serve: string;
  anchors: Record<string, string>;
  onChange: (segments: Segment[]) => void;
}

export function BodyComposer({ email, serve, anchors, onChange }: Props) {
  const segments = email.body_segments.length ? email.body_segments : [{ type: "text", value: "" } as Segment];
  const focus = useRef<{ idx: number; caret: number }>({ idx: 0, caret: 0 });
  const [render, setRender] = useState<{ text?: string; emissions?: Record<string, string>; error?: string }>({});
  const [menu, setMenu] = useState(false);
  const menuRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    let live = true;
    const id = setTimeout(async () => {
      try {
        const r = await api.render(segments, serve, anchors);
        if (live) setRender(r.ok ? { text: r.text, emissions: r.emissions } : { error: r.error });
      } catch {
        if (live) setRender({ error: "render unavailable" });
      }
    }, 150);
    return () => { live = false; clearTimeout(id); };
  }, [JSON.stringify(segments), serve, JSON.stringify(anchors)]);

  useEffect(() => {
    if (!menu) return;
    const onDoc = (e: MouseEvent) => { if (menuRef.current && !menuRef.current.contains(e.target as Node)) setMenu(false); };
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, [menu]);

  const setText = (idx: number, value: string) =>
    onChange(segments.map((s, i) => (i === idx && s.type === "text" ? { ...s, value } : s)));

  const insert = (piece: Segment) => {
    const { idx, caret } = focus.current;
    const seg = segments[idx];
    const next: Segment[] = [];
    segments.forEach((s, i) => {
      if (i === idx && s.type === "text") {
        next.push({ type: "text", value: s.value.slice(0, caret) });
        next.push(piece);
        next.push({ type: "text", value: s.value.slice(caret) });
      } else next.push(s);
    });
    if (!seg || seg.type !== "text") { next.push(piece); next.push({ type: "text", value: "" }); }
    onChange(coalesce(next));
    setMenu(false);
  };

  const replaceAt = (idx: number, piece: Segment) => onChange(segments.map((s, i) => (i === idx ? piece : s)));
  const removeAt = (idx: number) => {
    const next = segments.filter((_, i) => i !== idx);
    onChange(coalesce(next.length ? next : [{ type: "text", value: "" }]));
  };

  const onDrop = (e: React.DragEvent) => {
    e.preventDefault();
    const kind = e.dataTransfer.getData("chipKind") as ChipBase["kind"];
    if (kind) insert({ type: "chip", chip: emptyChip(kind) });
  };

  return (
    <div>
      <div
        onDragOver={(e) => e.preventDefault()}
        onDrop={onDrop}
        className="min-h-[92px] border border-edge bg-ink/40 p-3 leading-8"
      >
        <div className="flex flex-wrap items-center gap-x-0.5 gap-y-1">
          {segments.map((seg, idx) =>
            seg.type === "text" ? (
              <input
                key={idx}
                value={seg.value}
                placeholder={idx === 0 && segments.length === 1 ? "Write the email…" : ""}
                onChange={(e) => setText(idx, e.target.value)}
                onSelect={(e) => (focus.current = { idx, caret: (e.target as HTMLInputElement).selectionStart ?? 0 })}
                onKeyUp={(e) => (focus.current = { idx, caret: (e.target as HTMLInputElement).selectionStart ?? 0 })}
                onClick={(e) => (focus.current = { idx, caret: (e.target as HTMLInputElement).selectionStart ?? 0 })}
                className="min-w-[8px] bg-transparent text-[15px] text-white outline-none placeholder:text-muted/60"
                style={{ width: `${Math.max(seg.value.length, 1)}ch` }}
              />
            ) : seg.type === "chip" ? (
              <ChipPill
                key={idx}
                chip={seg.chip}
                onChange={(c: Chip) => replaceAt(idx, { type: "chip", chip: c })}
                onRemove={() => removeAt(idx)}
                serve={serve}
                anchors={anchors}
                reachableAnchors={email.reachable_anchors}
                allowEmit
              />
            ) : (
              <FactPill
                key={idx}
                seg={seg as FactSeg}
                onChange={(s) => replaceAt(idx, s)}
                onRemove={() => removeAt(idx)}
                reachableFacts={email.reachable_facts}
              />
            )
          )}
        </div>
      </div>

      <div ref={menuRef} className="relative mt-2">
        <button onClick={() => setMenu((m) => !m)}
          className="border border-edge bg-panel2 px-2.5 py-1 text-xs text-white/90 hover:border-accent hover:text-accent">
          + Insert date or value ▾
        </button>
        {menu && (
          <div className="absolute left-0 top-full z-50 mt-1 w-[280px] border border-edge bg-panel2 shadow-2xl animate-popin">
            {INSERTS.filter((o) => o.label !== "↩ Reuse a shared value" || email.reachable_facts.length).map((o) => (
              <button
                key={o.label}
                draggable={!!o.drag}
                onDragStart={(e) => o.drag && e.dataTransfer.setData("chipKind", o.drag)}
                onClick={() => insert(o.make(email.reachable_facts))}
                className="flex w-full items-center justify-between gap-3 border-b border-edge/60 px-3 py-2 text-left text-sm text-white/90 last:border-b-0 hover:bg-accent/10"
              >
                <span>{o.label}</span>
                <span className="text-[11px] text-muted">{o.hint}</span>
              </button>
            ))}
          </div>
        )}
      </div>

      <div className="mt-3 border border-edge/60 bg-panel/50 p-3">
        <div className="lbl mb-1">What the model reads · {serve}</div>
        {render.error
          ? <div className="text-sm text-warn">{render.error}</div>
          : <div className="text-[15px] text-white/90">{render.text || <span className="text-muted">…</span>}</div>}
        {render.emissions && Object.keys(render.emissions).length > 0 && (
          <div className="mt-2 flex flex-wrap gap-2">
            {Object.entries(render.emissions).map(([n, v]) => (
              <span key={n} className="border border-anchor/40 bg-anchor/10 px-2 py-0.5 text-xs text-anchor">📌 {n} = {v}</span>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

function coalesce(segs: Segment[]): Segment[] {
  const out: Segment[] = [];
  for (const s of segs) {
    const last = out[out.length - 1];
    if (s.type === "text" && last && last.type === "text") last.value += s.value;
    else out.push({ ...s });
  }
  if (out.length === 0 || out[0].type !== "text") out.unshift({ type: "text", value: "" });
  if (out[out.length - 1].type !== "text") out.push({ type: "text", value: "" });
  return out;
}
