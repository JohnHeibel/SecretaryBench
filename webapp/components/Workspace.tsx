"use client";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { usePathname, useRouter, useSearchParams } from "next/navigation";
import { deleteNode as apiDelete, fetchNodes, lintCorpus, oracleCorpus, saveNode } from "@/lib/api";
import { anchorsInCorpus, focusClosure, normalizeAnswer, withDerivedDateEdges } from "@/lib/grammar";
import { projectAtlasNode } from "@/lib/templates";
import type { CorpusNode, Edge, Email, LintResult, OracleResult } from "@/lib/types";
import Sidebar from "./Sidebar";
import EmailEditor from "./EmailEditor";
import DagCanvas from "./DagCanvas";
import ValidateBar from "./ValidateBar";

const EMPTY_EMAIL = (id: string): Email => ({
  id, from: "", to: "CEO", subject: "", body: "", depends_on: [], answer: { ops: [{ create: "", kind: "event", on: { eq: "" } }] },
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

  // Deep-link focus mode: ?node=<id> scopes one author to a single storyline (see
  // docs/AUTHORING_WALKTHROUGH.md). No param = the full "coordinator" view (assign storylines,
  // copy each author a focused link, run the whole-corpus check, export).
  const params = useSearchParams();
  const router = useRouter();
  const pathname = usePathname();
  const focusId = params.get("node");
  const focusEmail = params.get("email");
  const focused = focusId != null;

  // Clicking a storyline (or one of its emails, or a failing email in the bottom bar) leaves the
  // coordinator list and opens that storyline's own focused page (?node=<id>), so every author works
  // on a page scoped to one storyline. exitFocus drops the query to come back to the full list.
  const goFocus = useCallback((nodeId: string, emailId?: string) => {
    const q = new URLSearchParams({ node: nodeId }); if (emailId) q.set("email", emailId);
    router.push(`${pathname}?${q.toString()}`);
  }, [router, pathname]);
  const exitFocus = useCallback(() => router.push(pathname), [router, pathname]);

  // Pull the whole corpus from the store and seat it in state. Normalize each stored
  // answer into the verb-model shape at the data boundary, so a legacy/partial row can
  // never crash the editor downstream. In focus mode, open straight to the linked storyline.
  const reload = useCallback(async () => {
    const raw = await fetchNodes();
    const n = raw.map((node) => ({ ...node, emails: node.emails.map((e) => ({ ...e, answer: normalizeAnswer(e.answer) })) }));
    setNodes(n);
    if (focusId) {
      const hit = n.find((x) => x.id === focusId);
      setSelNode(hit ? focusId : null);
      // Honor ?email=<id> so a click that came from the coordinator list (or a bar "details" jump)
      // lands on the exact email, not just the storyline. Falls back to no selection if it's gone.
      setSelEmail(hit && focusEmail && hit.emails.some((e) => e.id === focusEmail) ? focusEmail : null);
    } else { setSelNode(n[0]?.id ?? null); setSelEmail(n[0]?.emails[0]?.id ?? null); }
    setLoaded(true);
  }, [focusId, focusEmail]);

  useEffect(() => { reload(); }, [reload]);

  // What lint/oracle run on: in focus mode just the open storyline (+ any it truly references, rare),
  // otherwise the whole corpus. Same real sb gate either way — focus mode just narrows the scope so one
  // author's keystrokes don't re-validate everyone else's in-flight work, and the bottom bar reflects
  // their storyline alone (in the common self-contained case the closure is exactly that one node).
  const validateNodes = useMemo(() => (focusId ? focusClosure(nodes, focusId) : nodes), [nodes, focusId]);

  // Re-lint (debounced) whenever the scoped corpus changes — the same gate the runner uses.
  useEffect(() => {
    if (!loaded) return;
    clearTimeout(lintTimer.current);
    lintTimer.current = setTimeout(() => { lintCorpus(validateNodes).then(setLint); }, 350);
  }, [validateNodes, loaded]);

  // Once it LINTS clean, prove it is also SOLVABLE: run the reference oracle (perfect secretary acting
  // from the answer key). Red here = an unsatisfiable answer key the linter cannot catch. We run it ONLY
  // in focus mode, on the one open storyline — in the coordinator list the bottom bar would show
  // oracle-red caused by some OTHER author's scenario, which reads as "my storyline is broken" when it
  // isn't. Scoped to the focused page, red here is always about THIS storyline. Gated on lint.ok too,
  // since the oracle needs a build-able corpus.
  useEffect(() => {
    if (!loaded) return;
    clearTimeout(oracleTimer.current);
    if (!focused || !lint?.ok) { setOracle(null); return; }
    oracleTimer.current = setTimeout(() => { oracleCorpus(validateNodes).then(setOracle); }, 350);
  }, [validateNodes, loaded, lint?.ok, focused]);

  // Anchors/variables offered to the editor are scoped to the SELECTED storyline only — an author
  // can't pick another storyline's published dates (cross-node references are an advanced, error-prone
  // case we don't surface). Lint/oracle still run on the full focus closure above, so a genuine
  // cross-node reference that already exists in the data is still validated, just not offered as a choice.
  const anchors = useMemo(() => anchorsInCorpus(nodes.filter((n) => n.id === selNode)), [nodes, selNode]);

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
  // pass. No-ops on collision so the caller can fall back. Only rows that actually changed are re-saved.
  const renameEmail = useCallback((oldId: string, newId: string): boolean => {
    if (!newId || newId === oldId || nodes.some((n) => n.emails.some((e) => e.id === newId))) return false;
    const fixEdge = (d: Edge): Edge => ({ ...d, email: d.email === oldId ? newId : d.email });
    const next = nodes.map((n) => ({ ...n,
      emails: n.emails.map((e) => ({ ...e, id: e.id === oldId ? newId : e.id, depends_on: e.depends_on.map(fixEdge) })),
      node_depends_on: n.node_depends_on?.map(fixEdge),
    }));
    setNodes(next);
    // Save only the rows that actually changed — a self-contained rename then touches just this storyline,
    // so one author's edit can't clobber a concurrent author's in-flight node with a stale snapshot.
    const before = new Map(nodes.map((n) => [n.id, JSON.stringify(n)]));
    for (const n of next) if (JSON.stringify(n) !== before.get(n.id)) saveNode(n).catch(() => {});
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

  // Load the fully-authored, oracle-solvable example storyline (Project Atlas) as a starting point. It
  // tours the headline constructs and shows the benchmark at scale (HOW_IT_WORKS.md).
  // Collisions on the node id get a -2, -3… suffix so loading it twice never clobbers an existing one.
  const addExample = useCallback(() => {
    const ids = new Set(nodes.map((n) => n.id));
    let id = "project_atlas", i = 2;
    while (ids.has(id)) id = `project_atlas-${i++}`;
    const node = projectAtlasNode(id);
    setNodes((prev) => [...prev, node]); setSelNode(node.id); setSelEmail(node.emails[0]?.id ?? null); persist(node);
  }, [nodes, persist]);

  const addEmail = useCallback((nodeId: string) => {
    const node = nodes.find((n) => n.id === nodeId); if (!node) return;
    let i = node.emails.length + 1; const ids = new Set(node.emails.map((e) => e.id));
    while (ids.has(`${nodeId}.e${i}`)) i++;
    const email = EMPTY_EMAIL(`${nodeId}.e${i}`);   // always addressed to the CEO (the assistant's boss)
    const cast = node.cast.CEO !== undefined ? node.cast : { ...node.cast, CEO: "you" };
    const updated = { ...node, cast, emails: [...node.emails, email] };
    updateNode(updated); setSelEmail(email.id);
  }, [nodes, updateNode]);

  const removeEmail = useCallback((nodeId: string, emailId: string) => {
    const node = nodes.find((n) => n.id === nodeId); if (!node) return;
    const label = node.emails.find((e) => e.id === emailId)?.subject?.trim() || emailId;
    if (!window.confirm(`Delete the email "${label}"? This can't be undone.`)) return;
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
    // Save only rows that changed (see renameEmail) — the renamed node has a new id so it always saves;
    // an untouched sibling storyline is left alone, protecting a concurrent author's edits.
    const before = new Map(nodes.map((n) => [n.id, JSON.stringify(n)]));
    for (const n of next) if (JSON.stringify(n) !== before.get(n.id)) saveNode(n).catch(() => {});
    apiDelete(oldId).catch(() => {});
    if (selNode === oldId) setSelNode(newId);
    if (selEmail && emailMap.has(selEmail)) setSelEmail(emailMap.get(selEmail)!);
    return true;
  }, [nodes, selNode, selEmail]);

  // Drag-connect from the DAG: make `targetId` depend on `prerequisiteId`. Same path the editor's
  // picker uses (new edges default to `static`; updateEmail runs withDerivedDateEdges so a needle still
  // auto-upgrades to `date`). No-ops if the edge already exists or they aren't in the named node.
  const addEdge = useCallback((nodeId: string, targetId: string, prerequisiteId: string) => {
    const node = nodes.find((n) => n.id === nodeId); if (!node) return;
    const target = node.emails.find((e) => e.id === targetId); if (!target) return;
    if (target.depends_on.some((d) => d.email === prerequisiteId)) return;
    updateEmail(nodeId, { ...target, depends_on: [...target.depends_on, { email: prerequisiteId, type: "static" }] });
  }, [nodes, updateEmail]);

  // Jump the editor straight to an email by id — used by the validation bar's "details" link so a
  // lint/oracle failure routes to exactly the email it's about. From the coordinator list that means
  // opening the owning storyline's focused page (where it gets fixed); in focus mode it just selects.
  const jumpToEmail = useCallback((emailId: string) => {
    const owner = nodes.find((n) => n.emails.some((e) => e.id === emailId)); if (!owner) return;
    if (focused) { setSelNode(owner.id); setSelEmail(emailId); setView("editor"); }
    else goFocus(owner.id, emailId);
  }, [nodes, focused, goFocus]);

  const removeNode = useCallback((nodeId: string) => {
    const count = nodes.find((n) => n.id === nodeId)?.emails.length ?? 0;
    const tail = count ? ` and its ${count} email${count === 1 ? "" : "s"}` : "";
    if (!window.confirm(`Delete the storyline "${nodeId}"${tail}? This can't be undone.`)) return;
    setNodes((prev) => prev.filter((n) => n.id !== nodeId));
    apiDelete(nodeId).catch(() => {});
    if (selNode === nodeId) { setSelNode(null); setSelEmail(null); }
  }, [nodes, selNode]);

  const node = nodes.find((n) => n.id === selNode) ?? null;
  const email = node?.emails.find((e) => e.id === selEmail) ?? null;
  const focusNode = focusId ? nodes.find((n) => n.id === focusId) ?? null : null;
  // In focus mode the sidebar, DAG, and "depends on" picker see only the author's own storyline.
  const viewNodes = focused ? (focusNode ? [focusNode] : []) : nodes;

  return (
    <div className="flex h-screen flex-col">
      <header className="flex flex-wrap items-center gap-3 border-b border-slate-800 bg-slate-900 px-4 py-2">
        <div className="flex min-w-0 items-center gap-3">
          {focused && <button onClick={exitFocus} title="Back to the full storyline list" className="shrink-0 rounded-md border border-slate-700 px-2 py-1 text-xs text-slate-300 hover:bg-slate-800">← All storylines</button>}
          <div className="min-w-0">
            <h1 className="text-sm font-semibold">SecretaryBench authoring</h1>
            <p className="hidden text-[11px] text-slate-500 md:block">{focused ? "You're editing one storyline. Write the email, say the perfect action, get the bottom bar green." : "Write an email, say the perfect action, then export once the bottom bar is green."}</p>
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
          <label className="flex items-center gap-1" title="Pretend this email arrived on this day, so the date chips show real dates while you write. It is not saved and does not change your export. The real run picks its own send date.">
            <span className="hidden md:inline">preview date</span>
            <input type="date" value={serveDate} onChange={(e) => setServeDate(e.target.value)} className="rounded border border-slate-700 bg-slate-800 px-2 py-0.5 text-slate-200" />
          </label>
          {!focused && <a href="/api/export" className="rounded-md border border-emerald-700/70 bg-emerald-600/10 px-3 py-1.5 font-medium text-emerald-100 hover:bg-emerald-600/20" title="Export when the bottom validation bar says Ready for export.">Export corpus</a>}
        </div>
      </header>

      <div className="flex min-h-0 flex-1">
        <Sidebar
          nodes={viewNodes} selNode={selNode} selEmail={selEmail} focused={focused}
          onSelectNode={focused ? (id) => { setSelNode(id); setSelEmail(null); } : (id) => goFocus(id)}
          onSelectEmail={focused ? (nid, eid) => { setSelNode(nid); setSelEmail(eid); } : (nid, eid) => goFocus(nid, eid)}
          onAddNode={addNode} onAddEmail={addEmail} onAddExample={addExample}
          onRemoveNode={removeNode} onRemoveEmail={removeEmail}
          lint={lint}
        />

        <main className="min-w-0 flex-1 overflow-auto">
          {view === "dag" ? (
            <DagCanvas nodes={viewNodes} lint={lint} serveDate={serveDate}
              onSelectEmail={focused ? (nid, eid) => { setSelNode(nid); setSelEmail(eid); setView("editor"); } : (nid, eid) => goFocus(nid, eid)}
              onAddEdge={addEdge} />
          ) : focused && !focusNode ? (
            <FocusNotFound focusId={focusId} loaded={loaded} />
          ) : node ? (
            <EmailEditor
              key={`${node.id}/${email?.id ?? "node"}`}
              node={node} email={email} allNodes={viewNodes} anchors={anchors} serveDate={serveDate}
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

      <ValidateBar lint={lint} oracle={oracle} nodes={validateNodes} showOracle={focused} onJump={jumpToEmail} />
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
        <p className="mt-3 text-xs text-slate-500">New here? <button onClick={onAddExample} className="text-sky-400 underline hover:text-sky-300">Load the Project Atlas example</button>. A compact product-launch saga with timed events, needles, a CEO-sent email, a no-action FYI, a timed deadline, a reschedule, and a cancel.</p>
      </div>
    </div>
  );
}

// Shown when ?node=<id> points at a storyline that isn't in the store. While the corpus is still
// loading we say nothing definitive — the id may just not be seated in state yet.
function FocusNotFound({ focusId, loaded }: { focusId: string | null; loaded: boolean }) {
  if (!loaded) return <div className="grid h-full place-items-center p-6 text-sm text-slate-500">Loading your storyline…</div>;
  return (
    <div className="grid h-full place-items-center p-6">
      <div className="max-w-md rounded-lg border border-slate-800 bg-slate-900/50 p-6 text-center">
        <p className="text-xs font-semibold uppercase tracking-wide text-amber-300">Storyline not found</p>
        <h2 className="mt-2 text-xl font-semibold text-slate-100">No storyline named <span className="font-mono text-amber-200">{focusId}</span>.</h2>
        <p className="mt-2 text-sm leading-6 text-slate-400">Double-check the link you were given, or ask whoever shared it to create this storyline and send you a fresh link.</p>
      </div>
    </div>
  );
}
