import { useEffect, useState } from "react";
import { api } from "../api";
import { BASE_LABELS, ORDINALS, UNITS, WEEKDAYS, WEEKDAY_LONG, emptyChip } from "../chipLabel";
import type { Chip, ChipBase, ResolveResult } from "../types";

const BASE_ORDER: ChipBase["kind"][] = [
  "next_weekday", "this_weekday", "serve", "anchor",
  "nth_weekday", "day_of_month", "month", "week_of", "raw",
];

interface Props {
  chip: Chip;
  onChange: (c: Chip) => void;
  serve: string;
  anchors: Record<string, string>;
  reachableAnchors: string[];
  allowEmit?: boolean;
}

export function ChipBuilder({ chip, onChange, serve, anchors, reachableAnchors, allowEmit }: Props) {
  const [preview, setPreview] = useState<ResolveResult | null>(null);

  useEffect(() => {
    let live = true;
    const id = setTimeout(async () => {
      try {
        const r = await api.resolve(chip, serve, anchors);
        if (live) setPreview(r);
      } catch {
        if (live) setPreview({ ok: false, error: "preview unavailable" });
      }
    }, 120);
    return () => { live = false; clearTimeout(id); };
  }, [JSON.stringify(chip), serve, JSON.stringify(anchors)]);

  const b = chip.base;
  const setBase = (patch: Partial<ChipBase>) => onChange({ ...chip, base: { ...b, ...patch } });
  const changeKind = (kind: ChipBase["kind"]) => {
    const fresh = emptyChip(kind);
    onChange({ ...fresh, offset: chip.offset, time: chip.time, emit_as: chip.emit_as });
  };

  const allowsOffset = b.kind !== "month" && b.kind !== "week_of" && b.kind !== "raw";
  const allowsTime = b.kind !== "month" && b.kind !== "week_of" && b.kind !== "raw";

  return (
    <div className="w-[320px] rounded-xl border border-edge bg-panel2 p-3 shadow-2xl animate-popin">
      <Row label="Kind">
        <select className={inp} value={b.kind} onChange={(e) => changeKind(e.target.value as any)}>
          {BASE_ORDER.map((k) => (
            <option key={k} value={k}>{BASE_LABELS[k]}</option>
          ))}
        </select>
      </Row>

      {(b.kind === "next_weekday" || b.kind === "this_weekday") && (
        <Row label="Weekday">
          <select className={inp} value={b.weekday} onChange={(e) => setBase({ weekday: e.target.value })}>
            {WEEKDAYS.map((w) => <option key={w} value={w}>{WEEKDAY_LONG[w]}</option>)}
          </select>
        </Row>
      )}

      {b.kind === "anchor" && (
        <Row label="Saved date">
          {reachableAnchors.length ? (
            <select className={inp} value={b.name} onChange={(e) => setBase({ name: e.target.value })}>
              <option value="">— pick a 📌 pin —</option>
              {reachableAnchors.map((n) => <option key={n} value={n}>📌 {n}</option>)}
            </select>
          ) : (
            <span className="text-xs text-muted">No upstream pins reach this email yet. Draw an arrow from the email that saves the date.</span>
          )}
        </Row>
      )}

      {b.kind === "nth_weekday" && (
        <>
          <Row label="Which">
            <select className={inp} value={String(b.n)} onChange={(e) => setBase({ n: e.target.value === "last" ? "last" : Number(e.target.value) })}>
              {ORDINALS.map((o) => <option key={String(o)} value={String(o)}>{o === "last" ? "last" : `${o}`}</option>)}
            </select>
          </Row>
          <Row label="Weekday">
            <select className={inp} value={b.weekday} onChange={(e) => setBase({ weekday: e.target.value })}>
              {WEEKDAYS.map((w) => <option key={w} value={w}>{WEEKDAY_LONG[w]}</option>)}
            </select>
          </Row>
          <MonthOffset value={b.month_offset ?? 0} onChange={(v) => setBase({ month_offset: v })} />
        </>
      )}

      {b.kind === "day_of_month" && (
        <>
          <Row label="Day">
            <input type="number" min={1} max={31} className={inp} value={b.day ?? 1}
              onChange={(e) => setBase({ day: Number(e.target.value) })} />
          </Row>
          <MonthOffset value={b.month_offset ?? 0} onChange={(v) => setBase({ month_offset: v })} />
        </>
      )}

      {b.kind === "month" && (
        <MonthOffset value={b.month_offset ?? 0} onChange={(v) => setBase({ month_offset: v })} />
      )}

      {b.kind === "week_of" && b.inner && (
        <div className="mt-2 rounded-lg border border-edge/60 p-2">
          <div className="mb-1 text-[11px] uppercase tracking-wide text-muted">Week containing…</div>
          <ChipBuilder chip={b.inner} onChange={(inner) => setBase({ inner })}
            serve={serve} anchors={anchors} reachableAnchors={reachableAnchors} />
        </div>
      )}

      {b.kind === "raw" && (
        <Row label="Token">
          <input className={`${inp} font-mono`} value={b.token}
            onChange={(e) => setBase({ token: e.target.value })} placeholder="e.g. nth:2,TUE,+1m" />
        </Row>
      )}

      {allowsOffset && (
        <div className="mt-2 flex items-end gap-2">
          <label className="flex items-center gap-1 text-xs text-muted">
            <input type="checkbox" checked={!!chip.offset}
              onChange={(e) => onChange({ ...chip, offset: e.target.checked ? { amount: 1, unit: "weeks" } : null })} />
            shift by
          </label>
          {chip.offset && (
            <>
              <input type="number" className={`${inp} w-16`} value={chip.offset.amount}
                onChange={(e) => onChange({ ...chip, offset: { ...chip.offset!, amount: Number(e.target.value) } })} />
              <select className={inp} value={chip.offset.unit}
                onChange={(e) => onChange({ ...chip, offset: { ...chip.offset!, unit: e.target.value } })}>
                {UNITS.map((u) => <option key={u.value} value={u.value}>{u.label}</option>)}
              </select>
            </>
          )}
        </div>
      )}

      {allowsTime && (
        <div className="mt-2 flex items-center gap-2">
          <label className="flex items-center gap-1 text-xs text-muted">
            <input type="checkbox" checked={!!chip.time}
              onChange={(e) => onChange({ ...chip, time: e.target.checked ? { hour: 9, minute: 0 } : null })} />
            at time
          </label>
          {chip.time && (
            <input type="time" className={inp}
              value={`${String(chip.time.hour).padStart(2, "0")}:${String(chip.time.minute).padStart(2, "0")}`}
              onChange={(e) => {
                const [h, m] = e.target.value.split(":").map(Number);
                onChange({ ...chip, time: { hour: h || 0, minute: m || 0 } });
              }} />
          )}
        </div>
      )}

      {allowEmit && (
        <div className="mt-2 rounded-lg border border-anchor/40 bg-anchor/10 p-2">
          <label className="flex items-center gap-2 text-xs">
            <input type="checkbox" checked={!!chip.emit_as}
              onChange={(e) => onChange({ ...chip, emit_as: e.target.checked ? "" : null })} />
            <span>📌 remember this date for later emails</span>
          </label>
          {chip.emit_as != null && (
            <input className={`${inp} mt-2`} placeholder="name it, e.g. signing" value={chip.emit_as}
              onChange={(e) => onChange({ ...chip, emit_as: e.target.value.replace(/[^A-Za-z0-9_]/g, "") })} />
          )}
        </div>
      )}

      <div className="mt-3 rounded-lg bg-ink/60 px-3 py-2 text-sm">
        {preview?.ok ? (
          <span className="text-ok">→ {preview.human}</span>
        ) : (
          <span className="text-warn">{preview?.error || "…"}</span>
        )}
      </div>
    </div>
  );
}

function MonthOffset({ value, onChange }: { value: number; onChange: (v: number) => void }) {
  return (
    <Row label="Month">
      <select className={inp} value={value} onChange={(e) => onChange(Number(e.target.value))}>
        <option value={0}>this month</option>
        <option value={1}>next month</option>
        <option value={2}>in 2 months</option>
        <option value={3}>in 3 months</option>
        <option value={-1}>last month</option>
      </select>
    </Row>
  );
}

function Row({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="mt-2 flex items-center gap-2">
      <span className="w-20 shrink-0 text-[11px] uppercase tracking-wide text-muted">{label}</span>
      <div className="flex-1">{children}</div>
    </div>
  );
}

const inp =
  "w-full rounded-md border border-edge bg-ink px-2 py-1 text-sm text-white outline-none focus:border-accent";
