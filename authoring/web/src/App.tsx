import { useStore } from "./store";
import { EmailEditor } from "./components/EmailEditor";
import { GraphCanvas } from "./components/GraphCanvas";
import { ValidationPanel } from "./components/ValidationPanel";

export default function App() {
  const { graph, selected, dirty, busy, oracle, save, runOracle, addThread, addEmail, reload } = useStore();
  const selectedThread = selected && graph ? graph.emails[selected]?.thread : null;
  const errorCount = graph?.errors.length ?? 0;

  return (
    <div className="flex h-full flex-col">
      <header className="flex items-center gap-3 border-b border-edge bg-panel px-4 py-2.5">
        <div className="flex items-center gap-2">
          <span className="text-lg">🗓️</span>
          <span className="font-semibold">SecretaryBench</span>
          <span className="text-muted">·</span>
          <span className="text-sm text-muted">Scenario Editor</span>
        </div>

        <div className="mx-2 h-5 w-px bg-edge" />

        <button onClick={addThread} className={btn}>+ Thread</button>
        <button
          onClick={() => selectedThread && addEmail(selectedThread)}
          disabled={!selectedThread}
          className={`${btn} disabled:opacity-40`}
          title={selectedThread ? `Add an email to ${selectedThread}` : "Select an email first"}
        >
          + Email
        </button>

        <div className="ml-auto flex items-center gap-2">
          {errorCount > 0 && (
            <span className="rounded-md border border-bad/50 bg-bad/10 px-2 py-1 text-xs text-bad">
              {errorCount} issue{errorCount === 1 ? "" : "s"}
            </span>
          )}
          {oracle && oracle.ok && (
            <span className={`rounded-md border px-2 py-1 text-xs ${
              oracle.score === 1 ? "border-ok/50 bg-ok/10 text-ok" : "border-warn/50 bg-warn/10 text-warn"}`}>
              oracle {Math.round((oracle.score || 0) * 100)}% ({oracle.passed}/{oracle.total})
            </span>
          )}
          <button onClick={runOracle} disabled={busy} className={btn}>Check solvable</button>
          <button onClick={reload} disabled={busy} className={btn}>Revert</button>
          <button
            onClick={save}
            disabled={busy || dirty.size === 0}
            className={`rounded-md px-3 py-1.5 text-sm font-medium transition disabled:opacity-40 ${
              dirty.size ? "bg-accent text-ink hover:brightness-110" : "bg-panel2 text-muted"}`}
          >
            {busy ? "Saving…" : dirty.size ? `Save ${dirty.size}` : "Saved"}
          </button>
        </div>
      </header>

      <div className="relative flex min-h-0 flex-1">
        <div className="relative min-w-0 flex-1">
          <GraphCanvas />
          <ValidationPanel />
          <div className="pointer-events-none absolute right-3 top-3 z-30 rounded-lg border border-edge bg-panel/80 px-3 py-1.5 text-[11px] text-muted backdrop-blur">
            drag <b className="text-white/80">handles</b> to link emails · click a node to edit · ⏰ = deadline edge
          </div>
        </div>
        {selected && <EmailEditor />}
      </div>
    </div>
  );
}

const btn =
  "rounded-md border border-edge bg-panel2 px-2.5 py-1.5 text-sm text-white/90 transition hover:border-accent hover:text-accent";
