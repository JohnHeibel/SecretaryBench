import { useEffect, useRef, useState } from "react";
import { api } from "../api";
import { BASE_LABELS, emptyChip } from "../chipLabel";
import type { Chip, ChipBase, EmailForm, Segment } from "../types";
import { ChipPill } from "./ChipPill";

// chip types offered in the quick palette (drag or click to insert)
const PALETTE: { kind: ChipBase["kind"]; icon: string; label: string }[] = [
  { kind: "next_weekday", icon: "📅", label: "Next weekday" },
  { kind: "serve", icon: "📅", label: "When it arrives" },
  { kind: "anchor", icon: "📌", label: "A saved date" },
  { kind: "day_of_month", icon: "📅", label: "Day of month" },
  { kind: "nth_weekday", icon: "📅", label: "Nth weekday" },
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

  const setText = (idx: number, value: string) => {
    const next = segments.map((s, i) => (i === idx && s.type === "text" ? { ...s, value } : s));
    onChange(next);
  };

  const insertChip = (chip: Chip) => {
    const { idx, caret } = focus.current;
    const seg = segments[idx];
    const next: Segment[] = [];
    segments.forEach((s, i) => {
      if (i === idx && s.type === "text") {
        const before = s.value.slice(0, caret);
        const after = s.value.slice(caret);
        next.push({ type: "text", value: before });
        next.push({ type: "chip", chip });
        next.push({ type: "text", value: after });
      } else {
        next.push(s);
      }
    });
    if (!seg || seg.type !== "text") {
      next.push({ type: "chip", chip });
      next.push({ type: "text", value: "" });
    }
    onChange(coalesce(next));
  };

  const updateChip = (idx: number, chip: Chip) =>
    onChange(segments.map((s, i) => (i === idx ? { type: "chip", chip } : s)));

  const removeChip = (idx: number) => {
    const next = segments.filter((_, i) => i !== idx);
    onChange(coalesce(next.length ? next : [{ type: "text", value: "" }]));
  };

  const onDrop = (e: React.DragEvent) => {
    e.preventDefault();
    const kind = e.dataTransfer.getData("chipKind") as ChipBase["kind"];
    if (kind) insertChip(emptyChip(kind));
  };

  return (
    <div>
      <div className="mb-2 flex flex-wrap gap-1.5">
        {PALETTE.map((p) => (
          <button
            key={p.kind}
            draggable
            onDragStart={(e) => e.dataTransfer.setData("chipKind", p.kind)}
            onClick={() => insertChip(emptyChip(p.kind))}
            title={`Insert: ${BASE_LABELS[p.kind]}`}
            className="cursor-grab rounded-md border border-edge bg-panel2 px-2 py-1 text-xs text-muted hover:border-accent hover:text-accent active:cursor-grabbing"
          >
            {p.icon} {p.label}
          </button>
        ))}
      </div>

      <div
        onDragOver={(e) => e.preventDefault()}
        onDrop={onDrop}
        className="min-h-[96px] rounded-lg border border-edge bg-ink/50 p-3 leading-8"
      >
        <div className="flex flex-wrap items-center gap-x-0.5 gap-y-1">
          {segments.map((seg, idx) =>
            seg.type === "text" ? (
              <input
                key={idx}
                value={seg.value}
                size={Math.max(seg.value.length, 1)}
                placeholder={idx === 0 && segments.length === 1 ? "Write the email…" : ""}
                onChange={(e) => setText(idx, e.target.value)}
                onSelect={(e) => (focus.current = { idx, caret: (e.target as HTMLInputElement).selectionStart ?? 0 })}
                onKeyUp={(e) => (focus.current = { idx, caret: (e.target as HTMLInputElement).selectionStart ?? 0 })}
                onClick={(e) => (focus.current = { idx, caret: (e.target as HTMLInputElement).selectionStart ?? 0 })}
                className="min-w-[8px] max-w-full bg-transparent text-[15px] text-white outline-none placeholder:text-muted/60"
                style={{ width: `${Math.max(seg.value.length, 1)}ch` }}
              />
            ) : (
              <ChipPill
                key={idx}
                chip={seg.chip}
                onChange={(c) => updateChip(idx, c)}
                onRemove={() => removeChip(idx)}
                serve={serve}
                anchors={anchors}
                reachableAnchors={email.reachable_anchors}
                allowEmit
              />
            )
          )}
        </div>
      </div>

      <div className="mt-2 rounded-lg border border-edge/60 bg-panel/60 p-3">
        <div className="mb-1 text-[11px] uppercase tracking-wide text-muted">What the model reads (on {serve})</div>
        {render.error ? (
          <div className="text-sm text-warn">{render.error}</div>
        ) : (
          <div className="text-[15px] text-white/90">{render.text || <span className="text-muted">…</span>}</div>
        )}
        {render.emissions && Object.keys(render.emissions).length > 0 && (
          <div className="mt-2 flex flex-wrap gap-2">
            {Object.entries(render.emissions).map(([n, v]) => (
              <span key={n} className="rounded border border-anchor/40 bg-anchor/10 px-2 py-0.5 text-xs text-anchor">
                📌 {n} = {v}
              </span>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

// merge adjacent text segments so the model never sees split words
function coalesce(segs: Segment[]): Segment[] {
  const out: Segment[] = [];
  for (const s of segs) {
    const last = out[out.length - 1];
    if (s.type === "text" && last && last.type === "text") {
      last.value += s.value;
    } else {
      out.push({ ...s });
    }
  }
  if (out.length === 0 || out[0].type !== "text") out.unshift({ type: "text", value: "" });
  if (out[out.length - 1].type !== "text") out.push({ type: "text", value: "" });
  return out;
}
