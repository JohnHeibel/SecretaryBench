"use client";
import { useState } from "react";
import type { CorpusNode, Edge, EdgeType, Email } from "@/lib/types";
import BodyEditor from "./BodyEditor";
import AnswerKeyBuilder from "./AnswerKeyBuilder";

interface Props {
  node: CorpusNode;
  email: Email | null;
  allNodes: CorpusNode[];
  anchors: string[];
  serveDate: string;
  onUpdateNode: (node: CorpusNode) => void;
  onRenameNode: (oldId: string, newId: string) => boolean;
  onUpdateEmail: (email: Email) => void;
}

function castKeys(node: CorpusNode): string[] {
  return Object.keys(node.cast ?? {});
}

export default function EmailEditor({ node, email, allNodes, anchors, serveDate, onUpdateNode, onRenameNode, onUpdateEmail }: Props) {
  return (
    <div className="mx-auto max-w-3xl space-y-5 p-5">
      <label className="flex items-center gap-2 text-xs font-semibold uppercase tracking-wide text-slate-400">
        Node id
        <input key={node.id} defaultValue={node.id} spellCheck={false}
          onKeyDown={(e) => { if (e.key === "Enter") (e.target as HTMLInputElement).blur(); }}
          onBlur={(e) => { if (!onRenameNode(node.id, e.target.value)) e.target.value = node.id; }}
          className="flex-1 rounded border border-slate-700 bg-slate-800 px-2 py-1 font-mono text-sm normal-case tracking-normal text-slate-100" />
      </label>
      <CastManager node={node} onUpdateNode={onUpdateNode} />
      {email ? (
        <EmailPanel node={node} email={email} allNodes={allNodes} anchors={anchors} serveDate={serveDate} onUpdateEmail={onUpdateEmail} />
      ) : (
        <p className="rounded-lg border border-dashed border-slate-800 px-4 py-8 text-center text-sm text-slate-500">
          Select an email in the sidebar, or <span className="text-slate-300">+ email</span> to add one to <code>{node.id}</code>.
        </p>
      )}
    </div>
  );
}

function CastManager({ node, onUpdateNode }: { node: CorpusNode; onUpdateNode: (n: CorpusNode) => void }) {
  const [open, setOpen] = useState(false);
  const entries = Object.entries(node.cast ?? {});

  function setCast(cast: Record<string, string>) { onUpdateNode({ ...node, cast }); }
  function rename(oldKey: string, newKey: string) {
    if (!newKey || newKey === oldKey || node.cast[newKey]) return;
    const next: Record<string, string> = {};
    for (const [k, v] of Object.entries(node.cast)) next[k === oldKey ? newKey : k] = v;
    setCast(next);
  }

  return (
    <div className="rounded-lg border border-slate-800 bg-slate-900/40">
      <button onClick={() => setOpen(!open)} className="flex w-full items-center justify-between px-3 py-2 text-xs font-semibold uppercase tracking-wide text-slate-400">
        <span>Cast — {node.id} <span className="font-normal lowercase text-slate-500">({entries.length} people)</span></span>
        <span>{open ? "▾" : "▸"}</span>
      </button>
      {open && (
        <div className="space-y-1.5 border-t border-slate-800 p-3">
          {entries.map(([key, name]) => (
            <div key={key} className="flex items-center gap-2">
              <input defaultValue={key} onBlur={(e) => rename(key, e.target.value.trim())}
                className="w-28 rounded border border-slate-700 bg-slate-800 px-2 py-1 font-mono text-xs text-slate-200" />
              <input value={name} onChange={(e) => setCast({ ...node.cast, [key]: e.target.value })}
                placeholder="Display name (Role)" className="flex-1 rounded border border-slate-700 bg-slate-800 px-2 py-1 text-xs text-slate-200" />
              <button onClick={() => { const c = { ...node.cast }; delete c[key]; setCast(c); }} className="px-1 text-xs text-slate-500 hover:text-rose-400">✕</button>
            </div>
          ))}
          <button onClick={() => setCast({ ...node.cast, [`PERSON_${entries.length + 1}`]: "" })}
            className="text-xs text-slate-500 hover:text-sky-300">+ person</button>
        </div>
      )}
    </div>
  );
}

function EmailPanel({ node, email, allNodes, anchors, serveDate, onUpdateEmail }: Omit<Props, "onUpdateNode" | "onRenameNode"> & { email: Email }) {
  const keys = castKeys(node);
  const toValue = Array.isArray(email.to) ? email.to[0] ?? "" : email.to;
  const set = (p: Partial<Email>) => onUpdateEmail({ ...email, ...p });

  return (
    <div className="space-y-4">
      <div className="flex items-center gap-2 text-xs text-slate-500">
        <span className="rounded bg-slate-800 px-2 py-0.5 font-mono text-slate-400">{email.id}</span>
      </div>

      <div className="grid grid-cols-2 gap-3">
        <label className="text-xs text-slate-400">From
          <select value={email.from} onChange={(e) => set({ from: e.target.value })}
            className="mt-1 w-full rounded border border-slate-700 bg-slate-800 px-2 py-1.5 text-sm text-slate-100">
            <option value="">—</option>
            {keys.map((k) => <option key={k} value={k}>{k} · {node.cast[k]}</option>)}
          </select>
        </label>
        <label className="text-xs text-slate-400">To
          <select value={toValue} onChange={(e) => set({ to: e.target.value })}
            className="mt-1 w-full rounded border border-slate-700 bg-slate-800 px-2 py-1.5 text-sm text-slate-100">
            <option value="">—</option>
            {keys.map((k) => <option key={k} value={k}>{k} · {node.cast[k]}</option>)}
          </select>
        </label>
      </div>

      <label className="block text-xs text-slate-400">Subject
        <input value={email.subject} onChange={(e) => set({ subject: e.target.value })}
          className="mt-1 w-full rounded border border-slate-700 bg-slate-800 px-3 py-1.5 text-sm text-slate-100" />
      </label>

      <BodyEditor body={email.body} anchors={anchors} serveDate={serveDate} onChange={(body) => set({ body })} />

      <DependencyPicker email={email} allNodes={allNodes} onChange={(depends_on) => set({ depends_on })} />

      <AnswerKeyBuilder answer={email.answer} anchors={anchors} serveDate={serveDate} onChange={(answer) => set({ answer })} />
    </div>
  );
}

function DependencyPicker({ email, allNodes, onChange }: { email: Email; allNodes: CorpusNode[]; onChange: (edges: Edge[]) => void }) {
  const others = allNodes.flatMap((n) => n.emails).filter((e) => e.id !== email.id);
  const chosen = new Set(email.depends_on.map((d) => d.email).filter(Boolean) as string[]);
  const [addId, setAddId] = useState("");

  function add(type: EdgeType) {
    if (!addId || chosen.has(addId)) return;
    onChange([...email.depends_on, { email: addId, type }]);
    setAddId("");
  }
  function setType(emailId: string, type: EdgeType) {
    onChange(email.depends_on.map((d) => d.email === emailId ? { ...d, type } : d));
  }

  return (
    <div className="rounded-lg border border-slate-800 bg-slate-900/40 p-3">
      <h3 className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-400">Depends on</h3>
      <div className="space-y-1.5">
        {email.depends_on.map((d, i) => (
          <div key={i} className="flex items-center gap-2 text-xs">
            <code className="flex-1 truncate rounded bg-slate-800 px-2 py-1 font-mono text-slate-300">{d.email ?? `@node:${d.node}`}</code>
            <select value={d.type} onChange={(e) => d.email && setType(d.email, e.target.value as EdgeType)}
              className="rounded border border-slate-700 bg-slate-800 px-2 py-1 text-slate-200">
              <option value="static">static (fact, no deadline)</option>
              <option value="date">date (carries a deadline)</option>
            </select>
            <button onClick={() => onChange(email.depends_on.filter((_, j) => j !== i))} className="px-1 text-slate-500 hover:text-rose-400">✕</button>
          </div>
        ))}
        <div className="flex items-center gap-2 text-xs">
          <select value={addId} onChange={(e) => setAddId(e.target.value)} className="flex-1 rounded border border-slate-700 bg-slate-800 px-2 py-1 text-slate-200">
            <option value="">add a prerequisite email…</option>
            {others.filter((e) => !chosen.has(e.id)).map((e) => <option key={e.id} value={e.id}>{e.id} — {e.subject}</option>)}
          </select>
          <button onClick={() => add("static")} disabled={!addId} className="rounded bg-slate-800 px-2 py-1 text-slate-200 hover:bg-slate-700 disabled:opacity-40">+ static</button>
          <button onClick={() => add("date")} disabled={!addId} className="rounded bg-slate-800 px-2 py-1 text-slate-200 hover:bg-slate-700 disabled:opacity-40">+ date</button>
        </div>
      </div>
    </div>
  );
}
