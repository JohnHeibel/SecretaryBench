import { useEffect, useRef, useState } from "react";

export interface FactSeg {
  type: "fact";
  name: string;
  value: string | null; // non-null => this email DEFINES the fact; null => references it
  token?: string;
}

interface Props {
  seg: FactSeg;
  onChange: (s: FactSeg) => void;
  onRemove: () => void;
  reachableFacts: string[];
}

// "90m" -> "90 minutes" for a friendly preview
function human(value: string): string {
  const m = value.trim().match(/^(\d+)\s*([hm])$/);
  if (!m) return value;
  return `${m[1]} ${m[2] === "h" ? "hour" : "minute"}${m[1] === "1" ? "" : "s"}`;
}

export function FactPill({ seg, onChange, onRemove, reachableFacts }: Props) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLSpanElement>(null);
  const defines = seg.value !== null;

  useEffect(() => {
    if (!open) return;
    const onDoc = (e: MouseEvent) => { if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false); };
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, [open]);

  const label = defines ? `${seg.name || "fact"} = ${human(seg.value || "")}` : seg.name || "pick a value";

  return (
    <span ref={ref} className="relative inline-flex align-baseline" contentEditable={false}>
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        className="inline-flex items-center gap-1 border border-fact/60 bg-fact/15 px-1.5 py-0.5 text-[13px] font-medium text-fact hover:brightness-125"
        title={defines ? "A shared value other emails can reuse" : "Reuses a shared value"}
      >
        📏 {label}
      </button>
      <button type="button" onClick={onRemove} className="ml-0.5 px-1 text-xs text-muted hover:text-bad">×</button>

      {open && (
        <span className="absolute left-0 top-[120%] z-50 block w-[260px] border border-edge bg-panel2 p-3 shadow-2xl animate-popin">
          <div className="lbl mb-1">{defines ? "This email defines a shared value" : "Reuse a shared value"}</div>
          {defines ? (
            <>
              <input
                className={inp}
                placeholder="name (e.g. client_meeting_len)"
                value={seg.name}
                onChange={(e) => onChange({ ...seg, name: e.target.value.replace(/[^A-Za-z0-9_]/g, "") })}
              />
              <input
                className={`${inp} mt-2`}
                placeholder="value (e.g. 90m)"
                value={seg.value || ""}
                onChange={(e) => onChange({ ...seg, value: e.target.value })}
              />
              <div className="mt-2 text-xs text-fact">→ reads as “{human(seg.value || "")}”</div>
            </>
          ) : (
            <select className={inp} value={seg.name} onChange={(e) => onChange({ ...seg, name: e.target.value })}>
              <option value="">— pick —</option>
              {reachableFacts.map((f) => <option key={f} value={f}>{f}</option>)}
            </select>
          )}
        </span>
      )}
    </span>
  );
}

const inp = "w-full border border-edge bg-ink px-2 py-1 text-sm text-white outline-none focus:border-fact";
