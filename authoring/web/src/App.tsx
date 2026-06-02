import { useStore } from "./store";
import { EmailEditor } from "./components/EmailEditor";
import { GraphCanvas } from "./components/GraphCanvas";
import { TabBar } from "./components/TabBar";
import { ValidationPanel } from "./components/ValidationPanel";

export default function App() {
  const { graph, selected, dirty, busy, oracle, save, runOracle, newEmail, reload, activeScenario } = useStore();
  const errorCount = graph?.errors.length ?? 0;
  const activeCount = graph
    ? graph.threads.filter((t) => t.scenario === activeScenario).reduce((n, t) => n + t.emails.length, 0)
    : 0;
  const isEmpty = !!graph && activeCount === 0;

  return (
    <div className="flex h-full flex-col bg-ink">
      <header className="flex items-center gap-3 border-b border-edge bg-panel px-4 py-2.5">
        <div className="flex items-center gap-2 font-semibold tracking-tight">
          <span className="grid h-6 w-6 place-items-center bg-accent text-ink">S</span>
          <span>SecretaryBench</span>
          <span className="text-sm font-normal text-muted">Scenario Editor</span>
        </div>

        <button onClick={newEmail} className="ml-3 bg-accent px-3 py-1.5 text-sm font-medium text-ink transition hover:brightness-110">
          + New email
        </button>

        <div className="ml-auto flex items-center gap-2">
          {errorCount > 0 && (
            <span className="border border-bad/50 bg-bad/10 px-2 py-1 text-xs text-bad">
              {errorCount} issue{errorCount === 1 ? "" : "s"}
            </span>
          )}
          {oracle && oracle.ok && (
            <span className={`border px-2 py-1 text-xs ${
              oracle.score === 1 ? "border-ok/50 bg-ok/10 text-ok" : "border-warn/50 bg-warn/10 text-warn"}`}>
              oracle {Math.round((oracle.score || 0) * 100)}%
            </span>
          )}
          <button onClick={runOracle} disabled={busy} className={btn}>Check solvable</button>
          <button onClick={reload} disabled={busy || dirty.size === 0} className={`${btn} disabled:opacity-40`}>Revert</button>
          <button
            onClick={save}
            disabled={busy || dirty.size === 0}
            className={`px-3 py-1.5 text-sm font-medium transition disabled:opacity-40 ${
              dirty.size ? "bg-ok text-ink hover:brightness-110" : "bg-panel2 text-muted"}`}
          >
            {busy ? "Saving…" : dirty.size ? `Save ${dirty.size}` : "Saved"}
          </button>
        </div>
      </header>

      <TabBar />

      <div className="relative flex min-h-0 flex-1">
        <div className="relative min-w-0 flex-1">
          <GraphCanvas />
          <ValidationPanel />
          {isEmpty && (
            <div className="pointer-events-none absolute inset-0 grid place-items-center">
              <div className="pointer-events-auto flex flex-col items-center gap-3 text-center">
                <div className="text-lg font-medium text-white/90">This scenario is empty</div>
                <div className="max-w-xs text-sm text-muted">
                  Each card is one email a secretary receives. Emails added here stay independent
                  from other tabs. Start your first conversation.
                </div>
                <button onClick={newEmail} className="bg-accent px-4 py-2 text-sm font-medium text-ink hover:brightness-110">
                  + New email
                </button>
              </div>
            </div>
          )}
        </div>
        {selected && <EmailEditor />}
      </div>
    </div>
  );
}

const btn =
  "border border-edge bg-panel2 px-2.5 py-1.5 text-sm text-white/90 transition hover:border-accent hover:text-accent";
