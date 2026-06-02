"use client";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { deleteNode as apiDelete, fetchNodes, lintCorpus, oracleCorpus, saveNode } from "@/lib/api";
import { anchorsInCorpus, normalizeAnswer } from "@/lib/grammar";
import type { CorpusNode, Email, LintResult, OracleResult } from "@/lib/types";
import Sidebar from "./Sidebar";
import EmailEditor from "./EmailEditor";
import DagCanvas from "./DagCanvas";
import ValidateBar from "./ValidateBar";

const EMPTY_EMAIL = (id: string): Email => ({
  id, from: "", to: "", subject: "", body: "", depends_on: [], answer: { ops: [] },
});

export default function Workspace() {
  const [nodes, setNodes] = useState<CorpusNode[]>([]);
  const [selNode, setSelNode] = useState<string | null>(null);
  const [selEmail, setSelEmail] = useState<string | null>(null);
  const [lint, setLint] = useState<LintResult | null>(null);
  const [oracle, setOracle] = useState<OracleResult | null>(null);
  const [serveDate, setServeDate] = useState("2026-06-01");
  const [view, setView] = useState<"editor" | "dag">("editor");
  const [loaded, setLoaded] = useState(false);

  const lintTimer = useRef<ReturnType<typeof setTimeout>>(undefined);
  const oracleTimer = useRef<ReturnType<typeof setTimeout>>(undefined);
  const saveTimers = useRef<Record<string, ReturnType<typeof setTimeout>>>({});

  useEffect(() => {
    fetchNodes().then((raw) => {
      // Normalize each stored answer into the verb-model shape at the data boundary, so a
      // legacy/partial row can never crash the editor downstream.
      const n = raw.map((node) => ({ ...node, emails: node.emails.map((e) => ({ ...e, answer: normalizeAnswer(e.answer) })) }));
      setNodes(n);
      setSelNode(n[0]?.id ?? null);
      setSelEmail(n[0]?.emails[0]?.id ?? null);
      setLoaded(true);
    });
  }, []);

  // Re-lint the whole corpus (debounced) whenever it changes — the same gate the runner uses.
  useEffect(() => {
    if (!loaded) return;
    clearTimeout(lintTimer.current);
    lintTimer.current = setTimeout(() => { lintCorpus(nodes).then(setLint); }, 350);
  }, [nodes, loaded]);

  // Once a corpus LINTS clean, prove it is also SOLVABLE: run the reference oracle
  // (perfect secretary acting from the answer key). Red here = an unsatisfiable answer
  // key the linter cannot catch. Gated on lint.ok since the oracle needs a build-able corpus.
  useEffect(() => {
    if (!loaded) return;
    clearTimeout(oracleTimer.current);
    if (!lint?.ok) { setOracle(null); return; }
    oracleTimer.current = setTimeout(() => { oracleCorpus(nodes).then(setOracle); }, 350);
  }, [nodes, loaded, lint?.ok]);

  const anchors = useMemo(() => anchorsInCorpus(nodes), [nodes]);

  const persist = useCallback((node: CorpusNode) => {
    clearTimeout(saveTimers.current[node.id]);
    saveTimers.current[node.id] = setTimeout(() => { saveNode(node).catch(() => {}); }, 500);
  }, []);

  const updateNode = useCallback((updated: CorpusNode) => {
    setNodes((prev) => prev.map((n) => (n.id === updated.id ? updated : n)));
    persist(updated);
  }, [persist]);

  const updateEmail = useCallback((nodeId: string, email: Email) => {
    setNodes((prev) => prev.map((n) => n.id === nodeId ? { ...n, emails: n.emails.map((e) => e.id === email.id ? email : e) } : n));
    const node = nodes.find((n) => n.id === nodeId);
    if (node) persist({ ...node, emails: node.emails.map((e) => e.id === email.id ? email : e) });
  }, [nodes, persist]);

  const addNode = useCallback(() => {
    const base = "node"; let i = 1; const ids = new Set(nodes.map((n) => n.id));
    while (ids.has(`${base}-${i}`)) i++;
    const node: CorpusNode = { id: `${base}-${i}`, cast: { CEO: "you" }, emails: [] };
    setNodes((prev) => [...prev, node]); setSelNode(node.id); setSelEmail(null); persist(node);
  }, [nodes, persist]);

  const addEmail = useCallback((nodeId: string) => {
    const node = nodes.find((n) => n.id === nodeId); if (!node) return;
    let i = node.emails.length + 1; const ids = new Set(node.emails.map((e) => e.id));
    while (ids.has(`${nodeId}.e${i}`)) i++;
    const email = EMPTY_EMAIL(`${nodeId}.e${i}`);
    const updated = { ...node, emails: [...node.emails, email] };
    updateNode(updated); setSelEmail(email.id);
  }, [nodes, updateNode]);

  const removeEmail = useCallback((nodeId: string, emailId: string) => {
    const node = nodes.find((n) => n.id === nodeId); if (!node) return;
    updateNode({ ...node, emails: node.emails.filter((e) => e.id !== emailId) });
    if (selEmail === emailId) setSelEmail(null);
  }, [nodes, updateNode, selEmail]);

  const removeNode = useCallback((nodeId: string) => {
    setNodes((prev) => prev.filter((n) => n.id !== nodeId));
    apiDelete(nodeId).catch(() => {});
    if (selNode === nodeId) { setSelNode(null); setSelEmail(null); }
  }, [selNode]);

  const node = nodes.find((n) => n.id === selNode) ?? null;
  const email = node?.emails.find((e) => e.id === selEmail) ?? null;

  return (
    <div className="flex h-screen flex-col">
      <header className="flex items-center justify-between border-b border-slate-800 bg-slate-900 px-4 py-2">
        <div className="flex items-center gap-3">
          <h1 className="text-sm font-semibold">SecretaryBench · Corpus Authoring</h1>
          <div className="flex overflow-hidden rounded-md border border-slate-700 text-xs">
            <button onClick={() => setView("editor")} className={`px-3 py-1 ${view === "editor" ? "bg-sky-600 text-white" : "text-slate-300 hover:bg-slate-800"}`}>Editor</button>
            <button onClick={() => setView("dag")} className={`px-3 py-1 ${view === "dag" ? "bg-sky-600 text-white" : "text-slate-300 hover:bg-slate-800"}`}>DAG</button>
          </div>
        </div>
        <div className="flex items-center gap-3 text-xs text-slate-400">
          <label className="flex items-center gap-1">preview serve date
            <input type="date" value={serveDate} onChange={(e) => setServeDate(e.target.value)} className="rounded border border-slate-700 bg-slate-800 px-2 py-0.5 text-slate-200" />
          </label>
          <a href="/api/export" className="rounded-md border border-slate-700 px-3 py-1 text-slate-200 hover:bg-slate-800">Export corpus ⬇</a>
        </div>
      </header>

      <div className="flex min-h-0 flex-1">
        <Sidebar
          nodes={nodes} selNode={selNode} selEmail={selEmail}
          onSelectNode={(id) => { setSelNode(id); setSelEmail(null); }}
          onSelectEmail={(nid, eid) => { setSelNode(nid); setSelEmail(eid); }}
          onAddNode={addNode} onAddEmail={addEmail}
          onRemoveNode={removeNode} onRemoveEmail={removeEmail}
          lint={lint}
        />

        <main className="min-w-0 flex-1 overflow-auto">
          {view === "dag" ? (
            <DagCanvas nodes={nodes} lint={lint} serveDate={serveDate}
              onSelectEmail={(nid, eid) => { setSelNode(nid); setSelEmail(eid); setView("editor"); }} />
          ) : node ? (
            <EmailEditor
              key={`${node.id}/${email?.id ?? "node"}`}
              node={node} email={email} allNodes={nodes} anchors={anchors} serveDate={serveDate}
              onUpdateNode={updateNode}
              onUpdateEmail={(e) => updateEmail(node.id, e)}
            />
          ) : (
            <div className="grid h-full place-items-center text-slate-500">Pick or create a node to start.</div>
          )}
        </main>
      </div>

      <ValidateBar lint={lint} oracle={oracle} />
    </div>
  );
}
