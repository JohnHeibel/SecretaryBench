import { useStore } from "../store";

export function ValidationPanel() {
  const { graph, oracle } = useStore();
  if (!graph) return null;
  const errors = graph.errors;
  if (errors.length === 0 && (!oracle || oracle.ok)) return null;

  return (
    <div className="absolute bottom-3 left-3 z-40 max-w-[640px] rounded-xl border border-bad/40 bg-panel/95 p-3 shadow-2xl backdrop-blur">
      <div className="mb-1 text-xs font-semibold uppercase tracking-wide text-bad">
        {errors.length > 0 ? `${errors.length} issue${errors.length === 1 ? "" : "s"} to fix` : "Oracle could not solve this corpus"}
      </div>
      <ul className="space-y-1 text-sm text-white/85">
        {errors.map((e, i) => <li key={i} className="font-mono text-xs">• {e}</li>)}
        {oracle && !oracle.ok && oracle.error && <li className="font-mono text-xs">• {oracle.error}</li>}
      </ul>
    </div>
  );
}
