"use client";
import type { LintResult, OracleResult } from "@/lib/types";

// The gate, mirrored from the runner: green only when the real sb.schema.lint passes.
// Once lint is green we ALSO show the oracle (satisfiability) result — green when the
// reference solver scores 1.0, red listing the email ids whose answer key is unsolvable.
export default function ValidateBar({ lint, oracle }: { lint: LintResult | null; oracle: OracleResult | null }) {
  if (!lint) {
    return <footer className="border-t border-slate-800 bg-slate-900 px-4 py-1.5 text-xs text-slate-500">validating…</footer>;
  }
  if (!lint.ok) {
    return (
      <footer className="border-t border-rose-900 bg-rose-950/50 px-4 py-1.5 text-xs text-rose-300">
        <span className="font-semibold">✗ won&apos;t pass the benchmark loader:</span> <span className="font-mono">{lint.error}</span>
      </footer>
    );
  }
  const s = lint.summary!;
  // lint passed — surface the oracle (satisfiability) status alongside it.
  if (oracle && oracle.error) {
    return (
      <footer className="border-t border-rose-900 bg-rose-950/50 px-4 py-1.5 text-xs text-rose-300">
        <span className="font-semibold">✗ oracle could not run:</span> <span className="font-mono">{oracle.error}</span>
      </footer>
    );
  }
  if (oracle && !oracle.ok) {
    return (
      <footer className="border-t border-rose-900 bg-rose-950/50 px-4 py-1.5 text-xs text-rose-300">
        <span className="font-semibold">✗ unsatisfiable answer key</span>
        <span className="ml-2 text-rose-400/80">oracle {oracle.passed}/{oracle.total} — no model could solve: <span className="font-mono">{oracle.failures!.join(", ")}</span></span>
      </footer>
    );
  }
  return (
    <footer className="flex items-center gap-3 border-t border-emerald-900 bg-emerald-950/50 px-4 py-1.5 text-xs text-emerald-300">
      <span className="font-semibold">✓ corpus valid</span>
      <span className="text-emerald-400/70">{s.nodes} nodes · {s.emails} emails · {s.anchors} anchors{oracle?.ok ? " · oracle solves 100%" : oracle ? "" : " · checking oracle…"} — ready to export &amp; run</span>
    </footer>
  );
}
