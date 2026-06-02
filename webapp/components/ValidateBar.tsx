"use client";
import type { LintResult, OracleResult } from "@/lib/types";

function friendlyError(error = ""): { title: string; body: string } {
  if (/op create needs a non-empty obligation name/.test(error)) return {
    title: "Finish the answer key",
    body: "Name the event or to-do you expect the assistant to create, like kickoff, signing, filing, or client review.",
  };
  if (/bad answer expression ''/.test(error) || /cannot parse base of ''/.test(error)) return {
    title: "Add the expected date",
    body: "The answer key has an empty date field. Use the date builder or type something like +5d, next:THU, or @signing+2w.",
  };
  if (/references undefined anchor/.test(error)) return {
    title: "That @anchor does not exist yet",
    body: "Create the earlier email that emits this anchor, or pick an existing anchor from the date builder.",
  };
  if (/serve-by window|has no 'date' dependency/.test(error)) return {
    title: "Add a date dependency",
    body: "This answer uses a date from an earlier email, so add that earlier email in Depends on as a date edge.",
  };
  if (/bad body token/.test(error)) return {
    title: "Fix a date token in the email body",
    body: "One token cannot be resolved. Rebuild it with the date builder so the email and answer key stay in sync.",
  };
  return { title: "Fix before export", body: error || "The corpus is not ready for the benchmark yet." };
}

// The gate, mirrored from the runner: green only when the real sb.schema.lint passes.
// Once lint is green we ALSO show the oracle (satisfiability) result — green when the
// reference solver scores 1.0, red listing the email ids whose answer key is unsolvable.
export default function ValidateBar({ lint, oracle }: { lint: LintResult | null; oracle: OracleResult | null }) {
  if (!lint) {
    return <footer className="border-t border-slate-800 bg-slate-900 px-4 py-2 text-xs text-slate-500">Checking the corpus…</footer>;
  }
  if (!lint.ok) {
    const msg = friendlyError(lint.error);
    return (
      <footer className="border-t border-rose-900 bg-rose-950/55 px-4 py-2 text-xs text-rose-200">
        <span className="font-semibold">{msg.title}:</span> <span>{msg.body}</span>
        <span className="ml-2 font-mono text-rose-300/70" title={lint.error}>details</span>
      </footer>
    );
  }
  const s = lint.summary!;
  if (s.nodes === 0) {
    return (
      <footer className="border-t border-slate-800 bg-slate-900 px-4 py-2 text-xs text-slate-400">
        Create one storyline to start. Validation and oracle checks become meaningful once it has emails.
      </footer>
    );
  }
  if (s.emails === 0) {
    return (
      <footer className="border-t border-slate-800 bg-slate-900 px-4 py-2 text-xs text-slate-400">
        Start by adding an email. Validation and oracle checks become meaningful once the corpus has at least one email.
      </footer>
    );
  }
  // lint passed — surface the oracle (satisfiability) status alongside it.
  if (oracle && oracle.error) {
    return (
      <footer className="border-t border-rose-900 bg-rose-950/55 px-4 py-2 text-xs text-rose-200">
        <span className="font-semibold">Oracle could not run:</span> <span className="font-mono">{oracle.error}</span>
      </footer>
    );
  }
  if (oracle && !oracle.ok) {
    if (oracle.total === 0) {
      return <footer className="border-t border-slate-800 bg-slate-900 px-4 py-2 text-xs text-slate-400">Add at least one email before running the oracle check.</footer>;
    }
    return (
      <footer className="border-t border-rose-900 bg-rose-950/55 px-4 py-2 text-xs text-rose-200">
        <span className="font-semibold">Answer key cannot be solved yet.</span>
        <span className="ml-2 text-rose-300/80">Oracle {oracle.passed}/{oracle.total}; fix: <span className="font-mono">{oracle.failures!.join(", ")}</span></span>
      </footer>
    );
  }
  return (
    <footer className="flex items-center gap-3 border-t border-emerald-900 bg-emerald-950/55 px-4 py-2 text-xs text-emerald-300">
      <span className="font-semibold">Ready for export</span>
      <span className="text-emerald-300/75">{s.nodes} nodes · {s.emails} emails · {s.anchors} anchors{oracle?.ok ? " · oracle solves 100%" : oracle ? "" : " · checking oracle…"}</span>
    </footer>
  );
}
