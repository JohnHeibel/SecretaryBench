"use client";
import type { CorpusNode, LintResult, OracleResult } from "@/lib/types";

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
    title: "Fix a date in the email body",
    body: "One date can’t be resolved. Rebuild it with the date builder so the email and answer key stay in sync.",
  };
  return { title: "Fix before export", body: error || "The corpus is not ready for the benchmark yet." };
}

// Find the email a lint error is about, so "details" can ALWAYS jump straight to it. Most CorpusError
// messages name it directly (`email '...'`). The op-level ones raised while parsing the answer key (blank
// or duplicate obligation name, missing/extra verb, missing predicate, …) name only the OBLIGATION, or
// nothing — so we resolve those against the live corpus: match a named obligation to the op that carries
// it, then fall back to the first structurally broken op (empty name / not exactly one verb). Returns null
// only when nothing can be pinpointed (e.g. a node/edge-level error), and then we show plain guidance.
function culpritEmailId(error: string, nodes: CorpusNode[]): string | null {
  const direct = error.match(/email '([^']+)'/)?.[1];
  if (direct) return direct;
  const verbName = (op: { create?: string; move?: string; cancel?: string }) => op.create ?? op.move ?? op.cancel;
  const named = error.match(/(?:op|obligation) '([^']+)'/)?.[1];
  if (named) {
    for (const n of nodes) for (const e of n.emails) for (const op of e.answer?.ops ?? []) if (verbName(op) === named) return e.id;
  }
  if (/non-empty obligation name|exactly one of create\/move\/cancel/.test(error)) {
    for (const n of nodes) for (const e of n.emails) for (const op of e.answer?.ops ?? []) {
      const verbs = (["create", "move", "cancel"] as const).filter((v) => v in op).length;
      if (!String(verbName(op) ?? "").trim() || verbs !== 1) return e.id;
    }
  }
  return null;
}

// A small clickable "→ id" pill that jumps the editor to the named email.
function JumpPill({ id, onJump }: { id: string; onJump?: (id: string) => void }) {
  if (!onJump) return <span className="font-mono text-rose-100">{id}</span>;
  return <button onClick={() => onJump(id)} title="Open this email" className="rounded bg-rose-900/70 px-1.5 py-0.5 font-mono text-rose-100 hover:bg-rose-800">→ {id}</button>;
}

// The gate, mirrored from the runner: green only when the real sb.schema.lint passes.
// Once lint is green we ALSO show the oracle (satisfiability) result — green when the
// reference solver scores 1.0, red naming each email whose answer key is unsolvable and why.
export default function ValidateBar({ lint, oracle, nodes, showOracle, onJump }: { lint: LintResult | null; oracle: OracleResult | null; nodes: CorpusNode[]; showOracle?: boolean; onJump?: (emailId: string) => void }) {
  if (!lint) {
    return <footer className="border-t border-slate-800 bg-slate-900 px-4 py-2 text-xs text-slate-500">Checking the corpus…</footer>;
  }
  if (!lint.ok) {
    const msg = friendlyError(lint.error);
    const where = culpritEmailId(lint.error ?? "", nodes);
    return (
      <footer className="border-t border-rose-900 bg-rose-950/55 px-4 py-2 text-xs text-rose-200">
        <span className="font-semibold">{msg.title}:</span> <span>{msg.body}</span>
        {/* "details" is the click target: when the error names an email, jump straight to it (raw error
            in the tooltip). When it can't (blank name/date), there's nowhere to point, so it stays a
            plain tooltip and the friendly body above is the guidance. */}
        {where && onJump
          ? <button onClick={() => onJump(where)} title={`${lint.error}\n\nClick to open ${where}`} className="ml-2 rounded bg-rose-900/70 px-1.5 py-0.5 font-mono text-rose-100 hover:bg-rose-800">details → {where}</button>
          : <span className="ml-2 font-mono text-rose-300/70" title={lint.error}>details</span>}
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
  // Lint passed. In the coordinator list (no storyline focused) we STOP here and just report
  // well-formed — the oracle (solvability) check runs per storyline on its focused page, so we never
  // show one author oracle-red that actually belongs to a different scenario. Open a storyline to see it.
  if (!showOracle) {
    return (
      <footer className="flex flex-wrap items-center gap-x-3 gap-y-1 border-t border-emerald-900 bg-emerald-950/55 px-4 py-2 text-xs text-emerald-300">
        <span className="font-semibold">Corpus is well-formed</span>
        <span className="text-emerald-300/75">{s.nodes} nodes · {s.emails} emails · {s.anchors} anchors</span>
        <span className="text-emerald-300/60">Open a storyline to run its oracle check.</span>
      </footer>
    );
  }
  // Focused on one storyline — surface the oracle (satisfiability) status alongside lint.
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
        <span className="font-semibold">Answer key cannot be solved yet</span>
        <span className="ml-2 text-rose-300/80">— the perfect secretary scored {oracle.passed}/{oracle.total}. Each line is what went wrong:</span>
        <ul className="mt-1 space-y-0.5">
          {oracle.failures!.map((f) => (
            <li key={f.id} className="flex items-baseline gap-1.5">
              <JumpPill id={f.id} onJump={onJump} />
              <span className="text-rose-200/90">{f.reason}</span>
            </li>
          ))}
        </ul>
      </footer>
    );
  }
  return (
    <footer className="flex items-center gap-3 border-t border-emerald-900 bg-emerald-950/55 px-4 py-2 text-xs text-emerald-300">
      <span className="font-semibold">This storyline is ready</span>
      <span className="text-emerald-300/75">{s.emails} emails · {s.anchors} anchors{oracle?.ok ? " · oracle solves 100%" : oracle ? "" : " · checking oracle…"}</span>
    </footer>
  );
}
