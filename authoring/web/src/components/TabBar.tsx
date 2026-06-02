import { useEffect, useRef, useState } from "react";
import { useStore } from "../store";
import { DEFAULT_SCENARIO } from "../types";

// The scenario switcher. Each tab is a self-contained scenario — nothing in one
// relates to anything in another (the linter enforces it), so switching tabs is a
// clean context switch. Double-click a tab to rename it; renaming a tab to an
// existing name merges the two.
export function TabBar() {
  const { graph, scenarios, activeScenario, setActiveScenario, createScenario, renameScenario, select } = useStore();

  const countByScenario: Record<string, number> = {};
  graph?.threads.forEach((t) => {
    countByScenario[t.scenario] = (countByScenario[t.scenario] || 0) + t.emails.length;
  });

  const [editing, setEditing] = useState<string | null>(null);
  const [draft, setDraft] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (editing) inputRef.current?.select();
  }, [editing]);

  const startEdit = (name: string) => {
    setEditing(name);
    setDraft(name === DEFAULT_SCENARIO ? "" : name);
  };
  const commit = () => {
    if (editing) {
      const next = draft.trim();
      if (next && next !== editing) renameScenario(editing, next);
    }
    setEditing(null);
  };

  const label = (name: string) => (name === DEFAULT_SCENARIO ? "Unsorted" : name);

  return (
    <div className="flex items-stretch gap-px overflow-x-auto border-b border-edge bg-ink scroll-thin">
      {scenarios.map((name) => {
        const active = name === activeScenario;
        const count = countByScenario[name] || 0;
        return (
          <button
            key={name}
            onClick={() => { setActiveScenario(name); select(null); }}
            onDoubleClick={() => startEdit(name)}
            title="Double-click to rename · rename onto another tab to merge"
            className={`flex items-center gap-2 whitespace-nowrap border-r border-edge px-3 py-1.5 text-sm transition ${
              active ? "bg-panel text-white" : "bg-ink text-muted hover:text-white/80"
            }`}
            style={active ? { boxShadow: "inset 0 -2px 0 0 #5b8cff" } : undefined}
          >
            {editing === name ? (
              <input
                ref={inputRef}
                value={draft}
                autoFocus
                onChange={(e) => setDraft(e.target.value)}
                onBlur={commit}
                onKeyDown={(e) => {
                  if (e.key === "Enter") commit();
                  if (e.key === "Escape") setEditing(null);
                }}
                onClick={(e) => e.stopPropagation()}
                className="w-28 border border-accent bg-ink px-1 text-sm text-white outline-none"
                placeholder="scenario name"
              />
            ) : (
              <>
                <span>{label(name)}</span>
                <span className={`text-[10px] ${active ? "text-muted" : "text-muted/70"}`}>{count}</span>
              </>
            )}
          </button>
        );
      })}
      <button
        onClick={() => { createScenario(); select(null); }}
        title="New scenario (independent — no relations to other tabs)"
        className="px-3 py-1.5 text-sm text-muted hover:text-accent"
      >
        + scenario
      </button>
    </div>
  );
}
