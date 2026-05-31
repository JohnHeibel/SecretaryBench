import { useEffect, useRef, useState } from "react";
import { chipLabel } from "../chipLabel";
import type { Chip } from "../types";
import { ChipBuilder } from "./ChipBuilder";

interface Props {
  chip: Chip;
  onChange: (c: Chip) => void;
  onRemove?: () => void;
  serve: string;
  anchors: Record<string, string>;
  reachableAnchors: string[];
  allowEmit?: boolean;
}

export function ChipPill({ chip, onChange, onRemove, serve, anchors, reachableAnchors, allowEmit }: Props) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLSpanElement>(null);

  useEffect(() => {
    if (!open) return;
    const onDoc = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, [open]);

  const isAnchorEmit = !!chip.emit_as;
  const tone = isAnchorEmit
    ? "border-anchor/60 bg-anchor/15 text-anchor"
    : "border-accent/50 bg-accent/15 text-accent";

  return (
    <span ref={ref} className="relative inline-flex align-baseline" contentEditable={false}>
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        className={`group inline-flex items-center gap-1 rounded-md border px-1.5 py-0.5 text-[13px] font-medium ${tone} hover:brightness-125`}
      >
        <span>{isAnchorEmit ? "📌 " : "📅 "}{isAnchorEmit ? `save “${chip.emit_as}” = ` : ""}{chipLabel(chip)}</span>
      </button>
      {onRemove && (
        <button
          type="button"
          onClick={onRemove}
          className="ml-0.5 rounded px-1 text-xs text-muted hover:text-bad"
          title="remove"
        >
          ×
        </button>
      )}
      {open && (
        <span className="absolute left-0 top-[120%] z-50 block">
          <ChipBuilder
            chip={chip}
            onChange={onChange}
            serve={serve}
            anchors={anchors}
            reachableAnchors={reachableAnchors}
            allowEmit={allowEmit}
          />
        </span>
      )}
    </span>
  );
}
