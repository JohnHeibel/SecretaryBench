"use client";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { deleteNode as apiDelete, fetchNodes, lintCorpus, oracleCorpus, saveNode } from "@/lib/api";
import { anchorsInCorpus, normalizeAnswer, withDerivedDateEdges } from "@/lib/grammar";
import { projectHeliosNode } from "@/lib/templates";
import type { CorpusNode, Edge, Email, LintResult, OracleResult } from "@/lib/types";
import Sidebar from "./Sidebar";
import EmailEditor from "./EmailEditor";
import DagCanvas from "./DagCanvas";
import ValidateBar from "./ValidateBar";

const EMPTY_EMAIL = (id: string): Email => ({
  id, from: "", to: "", subject: "", body: "", depends_on: [], answer: { ops: [{ create: "", kind: "event", on: { eq: "" } }] },
});

// A fresh email's id is `${node}.e${n}` — opaque. Once it has a subject we slug that into the id
// (`henderson.kickoff-reminder`) so the corpus JSON and dependency edges read like prose. We only
// auto-slug while the id is still in the default `.e<digits>` shape, so a deliberate id is never
// clobbered as the subject keeps changing.
const slugify = (s: string) => s.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 40).replace(/-+$/, "");
const isDefaultEmailId = (id: string, nodeId: string) => id.startsWith(`${nodeId}.e`) && /^\d+$/.test(id.slice(nodeId.length + 2));

export default function Workspace() {
  const [nodes, setNodes] = useState<CorpusNode[]>([]);
  const [selNode, setSelNode] = useState<string | null>(null);
  const [selEmail, setSelEmail] = useState<string | null>(null);
  const [lint, setLint] = useState<LintResult | null>(null);
  const [oracle, setOracle] = useState<OracleResult | null>(null);
  const [serveDate, setServeDate] = useState("2026-06-01");
  const [view, setView] = useState<"editor" | "dag">("editor");
  const [loaded, setLoaded] = useState(false);
  const [saveState, setSaveState] = useState<"idle" | "saving" | "saved" | "error">("idle");

  const lintTimer = useRef<ReturnType<typeof setTimeout>>(undefined);
  const oracleTimer = useRef<ReturnType<typeof setTimeout>>(undefined);
  const saveTimers = useRef<Record<string, ReturnType<typeof setTimeout>>>({});

  // Pull the whole corpus from the store and seat it in state. Normalize each stored
  // answer into the verb-model shape at the data boundary, so a legacy/partial row can
  // never crash the editor downstream.
  const reload = useCallback(async () => {
    const raw = await fetchNodes();
    const n = raw.map((node) => ({ ...node, emails: node.emails.map((e) => ({ ...e, answer: normalizeAnswer(e.answer) })) }));
    setNodes(n);
    setSelNode(n[0]?.id ?? null);
    setSelEmail(n[0]?.emails[0]?.id ?? null);
    setLoaded(true);
  }, []);

  useEffect(() => { reload(); }, [reload]);

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
    setSaveState("saving");
    saveTimers.current[node.id] = setTimeout(() => {
      saveNode(node).then(() => setSaveState("saved")).catch(() => setSaveState("error"));
    }, 500);
  }, []);

  const updateNode = useCallback((updated: CorpusNode) => {
    setNodes((prev) => prev.map((n) => (n.id === updated.id ? updated : n)));
    persist(updated);
  }, [persist]);

  const updateEmail = useCallback((nodeId: string, email: Email) => {
    // Auto-add the `date` dependency edge a needle (answer reusing an @anchor) requires, so authors
    // never hand-wire it — point a later email's answer at an earlier date and the link appears.
    const wired = withDerivedDateEdges(email, nodeId, nodes);
    setNodes((prev) => prev.map((n) => n.id === nodeId ? { ...n, emails: n.emails.map((e) => e.id === wired.id ? wired : e) } : n));
    const node = nodes.find((n) => n.id === nodeId);
    if (node) persist({ ...node, emails: node.emails.map((e) => e.id === wired.id ? wired : e) });
  }, [nodes, persist]);

  // Change an email's id (the corpus PK) and rewrite every depends_on edge across the whole corpus
  // that pointed at the old id, so cross-node references stay valid. Mirrors renameNode's edge-fix
  // pass. No-ops on collision so the caller can fall back. All touched nodes are re-saved.
  const renameEmail = useCallback((oldId: string, newId: string): boolean => {
    if (!newId || newId === oldId || nodes.some((n) => n.emails.some((e) => e.id === newId))) return false;
    const fixEdge = (d: Edge): Edge => ({ ...d, email: d.email === oldId ? newId : d.email });
    const next = nodes.map((n) => ({ ...n,
      emails: n.emails.map((e) => ({ ...e, id: e.id === oldId ? newId : e.id, depends_on: e.depends_on.map(fixEdge) })),
      node_depends_on: n.node_depends_on?.map(fixEdge),
    }));
    setNodes(next);
    for (const n of next) saveNode(n).catch(() => {});
    if (selEmail === oldId) setSelEmail(newId);
    return true;
  }, [nodes, selEmail]);

  // Fired when the subject field blurs. While the id is still the default `.e<n>`, derive it from
  // the subject; dedupe against the node's other emails by suffixing -2, -3… Once slugged the id is
  // no longer default, so it stays put through later subject edits.
  const autoSlugEmail = useCallback((nodeId: string, email: Email) => {
    if (!isDefaultEmailId(email.id, nodeId)) return;
    const base = slugify(email.subject); if (!base) return;
    const node = nodes.find((n) => n.id === nodeId); if (!node) return;
    const taken = new Set(node.emails.filter((e) => e.id !== email.id).map((e) => e.id));
    let id = `${nodeId}.${base}`, i = 2;
    while (taken.has(id)) id = `${nodeId}.${base}-${i++}`;
    renameEmail(email.id, id);
  }, [nodes, renameEmail]);

  const addNode = useCallback(() => {
    const base = "node"; let i = 1; const ids = new Set(nodes.map((n) => n.id));
    while (ids.has(`${base}-${i}`)) i++;
    const node: CorpusNode = { id: `${base}-${i}`, cast: { CEO: "you" }, emails: [] };
    setNodes((prev) => [...prev, node]); setSelNode(node.id); setSelEmail(null); persist(node);
  }, [nodes, persist]);

  // Load the worked "Project Helios" storyline (HOW_IT_WORKS.md) as a fully-authored, oracle-solvable
  // starting point — the fastest way to see a needle + reschedule + no-action thread end to end.
  const addExample = useCallback(() => {
    const ids = new Set(nodes.map((n) => n.id));
    let id = "project_helios", i = 2;
    while (ids.has(id)) id = `project_helios-${i++}`;
    const node = projectHeliosNode(id);
    setNodes((prev) => [...prev, node]); setSelNode(node.id); setSelEmail(node.emails[0]?.id ?? null); persist(node);
  }, [nodes, persist]);

  const addEmail = useCallback((nodeId: string) => {
    const node = nodes.find((n) => n.id === nodeId); if (!node) return;
    let i = node.emails.length + 1; const ids = new Set(node.emails.map((e) => e.id));
    while (ids.has(`${nodeId}.e${i}`)) i++;
    const email = { ...EMPTY_EMAIL(`${nodeId}.e${i}`), to: node.cast.CEO !== undefined ? "CEO" : "" };
    const updated = { ...node, emails: [...node.emails, email] };
    updateNode(updated); setSelEmail(email.id);
  }, [nodes, updateNode]);

  const removeEmail = useCallback((nodeId: string, emailId: string) => {
    const node = nodes.find((n) => n.id === nodeId); if (!node) return;
    updateNode({ ...node, emails: node.emails.filter((e) => e.id !== emailId) });
    if (selEmail === emailId) setSelEmail(null);
  }, [nodes, updateNode, selEmail]);

  // Rename = change the node id (the corpus PK). Re-prefix this node's emails that follow the
  // `${oldId}.` convention, then rewrite every depends_on / node_depends_on edge across the whole
  // corpus so cross-node references stay valid. Returns false (a no-op) on empty/collision so the
  // input can snap back. New rows are saved, then the old id is dropped from the store.
  const renameNode = useCallback((oldId: string, raw: string): boolean => {
    const newId = raw.trim();
    if (!newId || newId === oldId || nodes.some((n) => n.id === newId)) return false;
    const target = nodes.find((n) => n.id === oldId);
    const emailMap = new Map<string, string>();
    for (const e of target?.emails ?? []) if (e.id.startsWith(`${oldId}.`)) emailMap.set(e.id, `${newId}.${e.id.slice(oldId.length + 1)}`);
    const fixEdge = (d: Edge): Edge => ({ ...d, email: d.email ? emailMap.get(d.email) ?? d.email : d.email, node: d.node === oldId ? newId : d.node });
    const next = nodes.map((n) => {
      const base = { ...n, emails: n.emails.map((e) => ({ ...e, id: emailMap.get(e.id) ?? e.id, depends_on: e.depends_on.map(fixEdge) })), node_depends_on: n.node_depends_on?.map(fixEdge) };
      return n.id === oldId ? { ...base, id: newId } : base;
    });
    setNodes(next);
    for (const n of next) saveNode(n).catch(() => {});
    apiDelete(oldId).catch(() => {});
    if (selNode === oldId) setSelNode(newId);
    if (selEmail && emailMap.has(selEmail)) setSelEmail(emailMap.get(selEmail)!);
    return true;
  }, [nodes, selNode, selEmail]);

  const removeNode = useCallback((nodeId: string) => {
    setNodes((prev) => prev.filter((n) => n.id !== nodeId));
    apiDelete(nodeId).catch(() => {});
    if (selNode === nodeId) { setSelNode(null); setSelEmail(null); }
  }, [selNode]);

  const node = nodes.find((n) => n.id === selNode) ?? null;
  const email = node?.emails.find((e) => e.id === selEmail) ?? null;

  return (
    <div className="flex h-screen flex-col">
      <header className="flex flex-wrap items-center gap-3 border-b border-slate-800 bg-slate-900 px-4 py-2">
        <div className="flex min-w-0 items-center gap-3">
          <div className="min-w-0">
            <h1 className="text-sm font-semibold">SecretaryBench authoring</h1>
            <p className="hidden text-[11px] text-slate-500 md:block">Write emails, tell the grader the perfect action, then export when the bottom bar is green.</p>
          </div>
          <div className="flex overflow-hidden rounded-md border border-slate-700 text-xs">
            <button onClick={() => setView("editor")} className={`px-3 py-1 ${view === "editor" ? "bg-sky-600 text-white" : "text-slate-300 hover:bg-slate-800"}`}>Editor</button>
            <button onClick={() => setView("dag")} className={`px-3 py-1 ${view === "dag" ? "bg-sky-600 text-white" : "text-slate-300 hover:bg-slate-800"}`}>DAG</button>
          </div>
        </div>
        <div className="ml-auto flex items-center gap-3 text-xs text-slate-400">
          <span className={`hidden sm:inline ${saveState === "error" ? "text-rose-300" : saveState === "saving" ? "text-amber-300" : "text-slate-500"}`}>
            {saveState === "saving" ? "saving…" : saveState === "saved" ? "autosaved" : saveState === "error" ? "save failed" : "autosaves"}
          </span>
          <label className="flex items-center gap-1" title="A what-if: pretend this email arrived on this day, so the date chips show real dates while you author. It is NOT saved and does NOT change what you export — the real run picks its own send date.">
            <span className="hidden md:inline">preview date</span>
            <input type="date" value={serveDate} onChange={(e) => setServeDate(e.target.value)} className="rounded border border-slate-700 bg-slate-800 px-2 py-0.5 text-slate-200" />
          </label>
          <a href="/api/export" className="rounded-md border border-emerald-700/70 bg-emerald-600/10 px-3 py-1.5 font-medium text-emerald-100 hover:bg-emerald-600/20" title="Export when the bottom validation bar says Ready for export.">Export corpus</a>
        </div>
      </header>

      <div className="flex min-h-0 flex-1">
        <Sidebar
          nodes={nodes} selNode={selNode} selEmail={selEmail}
          onSelectNode={(id) => { setSelNode(id); setSelEmail(null); }}
          onSelectEmail={(nid, eid) => { setSelNode(nid); setSelEmail(eid); }}
          onAddNode={addNode} onAddEmail={addEmail} onAddExample={addExample}
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
              onUpdateNode={updateNode} onRenameNode={renameNode}
              onAddEmail={() => addEmail(node.id)}
              onUpdateEmail={(e) => updateEmail(node.id, e)}
              onAutoSlugEmail={(e) => autoSlugEmail(node.id, e)}
            />
          ) : (
            <StartEmpty onAddNode={addNode} onAddExample={addExample} />
          )}
        </main>
      </div>

      <ValidateBar lint={lint} oracle={oracle} />
    </div>
  );
}

function StartEmpty({ onAddNode, onAddExample }: { onAddNode: () => void; onAddExample: () => void }) {
  return (
    <div className="grid h-full place-items-center p-6">
      <div className="max-w-md rounded-lg border border-slate-800 bg-slate-900/50 p-6 text-center">
        <p className="text-xs font-semibold uppercase tracking-wide text-sky-300">First step</p>
        <h2 className="mt-2 text-xl font-semibold text-slate-100">Create one storyline.</h2>
        <p className="mt-2 text-sm leading-6 text-slate-400">A storyline is a group of related emails. Example: one deal, one client, one hiring process, or one office policy.</p>
        <button onClick={onAddNode} className="mt-5 rounded-md bg-sky-600 px-4 py-2 text-sm font-semibold text-white hover:bg-sky-500">Create first storyline</button>
        <p className="mt-3 text-xs text-slate-500">New here? <button onClick={onAddExample} className="text-sky-400 underline hover:text-sky-300">Load the Project Helios example</button> — a worked thread with a needle and a reschedule.</p>
      </div>
    </div>
  );
}
