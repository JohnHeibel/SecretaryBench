"use client";
import type { CorpusNode, LintResult } from "@/lib/types";
import { TIER_STYLE } from "./EmailEditor";

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
    <aside className="flex w-72 shrink-0 flex-col border-r border-slate-800 bg-slate-900/70">
      <div className="border-b border-slate-800 p-3">
        <div className="mb-2 flex items-center justify-between">
          <span className="text-xs font-semibold uppercase tracking-wide text-slate-400">Corpus</span>
          <span className="text-[11px] text-slate-500">{nodes.reduce((sum, n) => sum + n.emails.length, 0)} emails</span>
        </div>
        <button onClick={onAddNode} className="w-full rounded-md border border-sky-700/70 bg-sky-600/15 px-3 py-2 text-left text-sm font-medium text-sky-100 hover:bg-sky-600/25">
          + New storyline
          <span className="block text-[11px] font-normal text-sky-200/65">A storyline is a group of related emails.</span>
        </button>
      </div>
      <div className="min-h-0 flex-1 overflow-auto py-1">
        {nodes.map((node) => (
          <div key={node.id} className="px-2 py-1">
            <div className={`group flex items-center justify-between rounded-md px-2 py-1.5 ${selNode === node.id && !selEmail ? "bg-slate-800" : "hover:bg-slate-800/50"}`}>
              <button onClick={() => onSelectNode(node.id)} className="min-w-0 flex-1 truncate text-left text-sm font-medium text-slate-200">
                {node.id} <span className="text-xs text-slate-500">({node.emails.length})</span>
              </button>
              <button onClick={() => onRemoveNode(node.id)} className="hidden px-1 text-xs text-slate-500 hover:text-rose-400 group-hover:block">✕</button>
            </div>
            <div className="ml-3 border-l border-slate-800 pl-2">
              {node.emails.map((email) => (
                <div key={email.id} className={`group flex items-center justify-between rounded-md px-2 py-1 ${selEmail === email.id ? "bg-sky-600/20 text-sky-200" : "hover:bg-slate-800/50"}`}>
                  <button onClick={() => onSelectEmail(node.id, email.id)} className="min-w-0 flex-1 truncate text-left text-xs text-slate-300">
                    {email.subject || email.id}
                  </button>
                  {email.tier && <span className={`rounded px-1 text-[10px] font-medium ${TIER_STYLE[email.tier]}`}>{email.tier}</span>}
                  <button onClick={() => onRemoveEmail(node.id, email.id)} className="hidden px-1 text-xs text-slate-500 hover:text-rose-400 group-hover:block">✕</button>
                </div>
              ))}
              <button onClick={() => onAddEmail(node.id)} className="mt-1 w-full rounded border border-dashed border-slate-700 px-2 py-1.5 text-left text-xs text-slate-400 hover:border-sky-600 hover:text-sky-200">+ Add email</button>
            </div>
          </div>
        ))}
        {nodes.length === 0 && (
          <div className="mx-3 mt-3 rounded-md border border-dashed border-slate-700 p-3 text-xs text-slate-400">
            <p className="font-medium text-slate-300">Nothing here yet.</p>
            <p className="mt-1">Start with one storyline, then add the emails inside it.</p>
          </div>
        )}
      </div>
    </aside>
  );
}
