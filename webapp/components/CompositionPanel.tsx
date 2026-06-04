"use client";
import { useState } from "react";
import type { CorpusNode, Tier } from "@/lib/types";
import { type Band, MIN_ASSESS, ROLE_META, computeComposition } from "@/lib/composition";

// Live read-out of the corpus mix against the guide's target bands (AUTHORING_GUIDE §2–3).
// Sits at the top of the sidebar so an author sees, while writing, whether the inbox is
// drifting toward all-action / no-junk. Judgement (green ✓ / amber !) is withheld until there
// are enough emails for the percentages to mean anything (MIN_ASSESS).

const TIER_DOT: Record<Tier, string> = { T1: "bg-emerald-400", T2: "bg-amber-400", T3: "bg-rose-400" };
const fmtBand = (b: Band) => `${b.lo}–${b.hi}%`;

export default function CompositionPanel({ nodes }: { nodes: CorpusNode[] }) {
  const [open, setOpen] = useState(true);
  const c = computeComposition(nodes);
  const assess = c.total >= MIN_ASSESS;
  const offBand = assess ? c.roles.filter((r) => !r.inBand).length : 0;

  return (
    <div className="border-b border-slate-800 px-3 py-2">
      <button onClick={() => setOpen(!open)} className="flex w-full items-center justify-between text-xs font-semibold uppercase tracking-wide text-slate-400">
        <span>Corpus mix</span>
        <span className="flex items-center gap-1.5 font-normal normal-case">
          {assess && (offBand === 0
            ? <span className="text-emerald-400">on target</span>
            : <span className="text-amber-400">{offBand} off target</span>)}
          <span className="text-slate-500">{open ? "▾" : "▸"}</span>
        </span>
      </button>

      {open && (c.total === 0 ? (
        <p className="mt-2 text-[11px] leading-snug text-slate-500">No emails yet. The guide aims for roughly <span className="text-slate-400">80% action</span>, <span className="text-slate-400">12% FYI</span>, <span className="text-slate-400">7% junk</span>.</p>
      ) : (
        <div className="mt-2 space-y-2">
          <StackBar roles={c.roles} />
          <div className="space-y-1">
            {c.roles.map((r) => (
              <Row key={r.role} label={ROLE_META[r.role].label} dot={ROLE_META[r.role].dot}
                count={r.count} total={c.total} pct={r.pct} target={r.target} inBand={r.inBand} assess={assess} />
            ))}
          </div>

          {c.actionTotal > 0 && (
            <div className="mt-1.5 border-t border-slate-800 pt-1.5">
              <p className="mb-1 text-[10px] uppercase tracking-wide text-slate-500">Action difficulty <span className="font-normal lowercase text-slate-600">({c.actionTotal} of {c.total})</span></p>
              <div className="space-y-1">
                {c.tiers.map((t) => (
                  <Row key={t.tier} label={t.tier} dot={TIER_DOT[t.tier]}
                    count={t.count} total={c.actionTotal} pct={t.pct} target={t.target} inBand={t.inBand} assess={c.actionTotal >= MIN_ASSESS} />
                ))}
              </div>
              {c.untaggedAction > 0 && (
                <p className="mt-1 text-[11px] leading-snug text-amber-400">{c.untaggedAction} action {c.untaggedAction === 1 ? "email has" : "emails have"} no difficulty — tag T1/T2/T3.</p>
              )}
            </div>
          )}

          {!assess && <p className="text-[10px] leading-snug text-slate-600">Targets are judged once you have ~{MIN_ASSESS}+ emails.</p>}
          <a href="/guide" className="block text-[10px] text-slate-600 hover:text-sky-400">guide targets: §2 types · §3 tiers</a>
        </div>
      ))}
    </div>
  );
}

// A single stacked bar of the three role shares — the at-a-glance shape of the inbox.
function StackBar({ roles }: { roles: { role: keyof typeof ROLE_META; pct: number }[] }) {
  return (
    <div className="flex h-2 overflow-hidden rounded-full bg-slate-800" title="action · no-action · junk">
      {roles.map((r) => r.pct > 0 && <div key={r.role} className={ROLE_META[r.role].bar} style={{ width: `${r.pct}%` }} />)}
    </div>
  );
}

function Row({ label, dot, count, total, pct, target, inBand, assess }: {
  label: string; dot: string; count: number; total: number; pct: number; target: Band; inBand: boolean; assess: boolean;
}) {
  const judge = assess ? (inBand ? "text-emerald-400" : "text-amber-400") : "text-slate-600";
  const mark = assess ? (inBand ? "✓" : "!") : "";
  return (
    <div className="flex items-center gap-1.5 text-[11px]" title={`${count} of ${total}`}>
      <span className={`h-2 w-2 shrink-0 rounded-full ${dot}`} />
      <span className="text-slate-300">{label}</span>
      <span className="ml-auto tabular-nums text-slate-400">{Math.round(pct)}%</span>
      <span className="w-12 text-right tabular-nums text-slate-600">{fmtBand(target)}</span>
      <span className={`w-2.5 text-center ${judge}`}>{mark}</span>
    </div>
  );
}
