"use client";
import type { CorpusNode, LintResult } from "@/lib/types";

interface Props {
  nodes: CorpusNode[];
  selNode: string | null;
  selEmail: string | null;
  lint: LintResult | null;
  onSelectNode: (id: string) => void;
  onSelectEmail: (nodeId: string, emailId: string) => void;
  onAddNode: () => void;
  onAddEmail: (nodeId: string) => void;
  onRemoveNode: (nodeId: string) => void;
  onRemoveEmail: (nodeId: string, emailId: string) => void;
}

export default function Sidebar({ nodes, selNode, selEmail, onSelectNode, onSelectEmail, onAddNode, onAddEmail, onRemoveNode, onRemoveEmail }: Props) {
  return (
    <aside className="flex w-64 shrink-0 flex-col border-r border-slate-800 bg-slate-900/50">
      <div className="flex items-center justify-between border-b border-slate-800 px-3 py-2">
        <span className="text-xs font-semibold uppercase tracking-wide text-slate-400">Corpus</span>
        <button onClick={onAddNode} className="rounded bg-slate-800 px-2 py-0.5 text-xs text-slate-200 hover:bg-slate-700">+ node</button>
      </div>
      <div className="min-h-0 flex-1 overflow-auto py-1">
        {nodes.map((node) => (
          <div key={node.id} className="px-1">
            <div className={`group flex items-center justify-between rounded px-2 py-1 ${selNode === node.id && !selEmail ? "bg-slate-800" : "hover:bg-slate-800/50"}`}>
              <button onClick={() => onSelectNode(node.id)} className="min-w-0 flex-1 truncate text-left text-sm font-medium text-slate-200">
                {node.id} <span className="text-xs text-slate-500">({node.emails.length})</span>
              </button>
              <button onClick={() => onRemoveNode(node.id)} className="hidden px-1 text-xs text-slate-500 hover:text-rose-400 group-hover:block">✕</button>
            </div>
            <div className="ml-3 border-l border-slate-800 pl-1">
              {node.emails.map((email) => (
                <div key={email.id} className={`group flex items-center justify-between rounded px-2 py-0.5 ${selEmail === email.id ? "bg-sky-600/20 text-sky-200" : "hover:bg-slate-800/50"}`}>
                  <button onClick={() => onSelectEmail(node.id, email.id)} className="min-w-0 flex-1 truncate text-left text-xs text-slate-300">
                    {email.subject || email.id}
                  </button>
                  <button onClick={() => onRemoveEmail(node.id, email.id)} className="hidden px-1 text-xs text-slate-500 hover:text-rose-400 group-hover:block">✕</button>
                </div>
              ))}
              <button onClick={() => onAddEmail(node.id)} className="mt-0.5 px-2 py-0.5 text-xs text-slate-500 hover:text-sky-300">+ email</button>
            </div>
          </div>
        ))}
        {nodes.length === 0 && <p className="px-3 py-4 text-xs text-slate-500">No nodes yet. Create one to begin.</p>}
      </div>
    </aside>
  );
}
