import {
  Background,
  BackgroundVariant,
  Controls,
  MarkerType,
  MiniMap,
  Panel,
  ReactFlow,
  useEdgesState,
  useNodesState,
  type Connection,
  type Edge as RFEdge,
  type Node as RFNode,
  type NodeChange,
} from "@xyflow/react";
import "@xyflow/react/dist/style.css";
import { useCallback, useEffect, useMemo, useState } from "react";
import { chipLabel } from "../chipLabel";
import { useStore } from "../store";
import type { AnswerForm, Chip, Graph } from "../types";
import { EmailNode, type EmailNodeData } from "./EmailNode";

const THREAD_COLORS = ["#6aa6ff", "#34d399", "#fbbf24", "#fb7185", "#c084fc", "#22d3ee", "#f472b6", "#a3e635"];
const nodeTypes = { email: EmailNode };

// --- reading the answer for what it DEPENDS ON -----------------------------

// Every place an answer can reference a shared scalar by name, e.g. a
// `duration` of "@client_meeting_len" or a `count` of "@headcount".
function factRefsInAnswer(answer: AnswerForm): Set<string> {
  const out = new Set<string>();
  const scan = (v: unknown) => {
    if (typeof v !== "string") return;
    for (const m of v.matchAll(/@([A-Za-z_]\w*)/g)) out.add(m[1]);
  };
  answer.expect.forEach((ex) => {
    scan(ex.duration);
    scan(ex.count);
  });
  return out;
}

function anchorNamesInChip(chip: Chip | undefined, out: Set<string>) {
  if (!chip) return;
  if (chip.base.kind === "anchor" && chip.base.name) out.add(chip.base.name);
  if (chip.base.kind === "week_of") anchorNamesInChip(chip.base.inner, out);
}

// Saved-date ("anchor") names this answer is graded against — the email that
// EMITS each is a real prerequisite, expressed as a `date` dependency edge.
function anchorRefsInAnswer(answer: AnswerForm): Set<string> {
  const out = new Set<string>();
  answer.expect.forEach((ex) => {
    const w = ex.when;
    if (!w) return;
    anchorNamesInChip(w.chip, out);
    anchorNamesInChip(w.avoid_chip, out);
    (w.chips || []).forEach((c) => anchorNamesInChip(c, out));
  });
  return out;
}

function bodyPreview(g: Graph, id: string): string {
  const segs = g.emails[id]?.body_segments || [];
  return segs
    .map((s) =>
      s.type === "text" ? s.value : s.type === "chip" ? `「${chipLabel(s.chip)}」` : `「${s.name}」`
    )
    .join("");
}

function build(
  graph: Graph,
  selected: string | null,
  oracle: ReturnType<typeof useStore>["oracle"],
  onFollowUp: (id: string) => void,
  activeScenario: string
) {
  const colorOf: Record<string, string> = {};
  graph.threads.forEach((t, i) => (colorOf[t.id] = THREAD_COLORS[i % THREAD_COLORS.length]));
  // One scenario per tab: only show its threads. Relations can't cross scenarios,
  // so every edge is guaranteed to stay inside the active tab.
  const scenarioOf: Record<string, string> = {};
  graph.threads.forEach((t) => (scenarioOf[t.id] = t.scenario));
  const inScenario = (eid: string) => {
    const e = graph.emails[eid];
    return !!e && scenarioOf[e.thread] === activeScenario;
  };
  const emitsBy: Record<string, string[]> = {};
  Object.entries(graph.emission_map).forEach(([name, eid]) => {
    (emitsBy[eid] ||= []).push(name);
  });
  const factsBy: Record<string, string[]> = {};
  Object.entries(graph.fact_map || {}).forEach(([name, eid]) => {
    (factsBy[eid] ||= []).push(name);
  });

  const nodes: RFNode<EmailNodeData>[] = Object.values(graph.emails).filter((e) => inScenario(e.id)).map((e) => ({
    id: e.id,
    type: "email",
    position: { x: e.x ?? 80, y: e.y ?? 80 },
    data: {
      email: e,
      color: colorOf[e.thread] || "#5b8cff",
      selected: selected === e.id,
      oracle: oracle?.results?.[e.id],
      emitsAnchors: emitsBy[e.id] || [],
      definesFacts: factsBy[e.id] || [],
      bodyPreview: bodyPreview(graph, e.id),
      onFollowUp,
    },
  }));

  const edges: RFEdge[] = [];
  const threadOf = (eid: string) => graph.emails[eid]?.thread;
  const pairKey = (a: string, b: string) => `${a}|${b}`;
  // thread pairs already joined by a fine-grained edge — so we don't also draw a
  // redundant whole-thread (node-level) edge on top of them.
  const linkedThreadPairs = new Set<string>();
  const noteCrossThread = (srcEid: string, tgtEid: string) => {
    const a = threadOf(srcEid);
    const b = threadOf(tgtEid);
    if (a && b && a !== b) linkedThreadPairs.add(pairKey(a, b));
  };

  Object.values(graph.emails).filter((e) => inScenario(e.id)).forEach((e) => {
    // 1) Explicit ordering dependencies the author drew.
    e.depends_on.forEach((d) => {
      if (!d.email || !graph.emails[d.email] || !inScenario(d.email)) return;
      const isDate = d.type === "date";
      noteCrossThread(d.email, e.id);
      edges.push({
        id: `dep:${d.email}->${e.id}`,
        source: d.email,
        target: e.id,
        animated: isDate,
        label: isDate ? "deadline" : "after",
        labelStyle: { fill: isDate ? "#b98cff" : "#888f9c", fontSize: 10 },
        labelBgStyle: { fill: "#14171c" },
        labelBgPadding: [4, 2] as [number, number],
        markerEnd: { type: MarkerType.ArrowClosed, color: isDate ? "#b98cff" : "#5a6373", width: 16, height: 16 },
        style: { stroke: isDate ? "#b98cff" : "#39414e", strokeWidth: 1.5, strokeDasharray: isDate ? "5 4" : undefined },
      });
    });

    // 2) VALUE edges: this email's answer is graded against a shared fact (e.g.
    // duration "@client_meeting_len"), so it depends on the email that DEFINES
    // that value. No `depends_on` entry backs this, so draw it explicitly or the
    // defining email looks unconnected. Promoted (thicker, on top) because it's
    // a real grading constraint that often spans far-apart threads.
    factRefsInAnswer(e.answer).forEach((name) => {
      const src = graph.fact_map?.[name];
      if (!src || src === e.id || !graph.emails[src] || !inScenario(src)) return;
      noteCrossThread(src, e.id);
      edges.push({
        id: `fact:${name}:${src}->${e.id}`,
        source: src,
        target: e.id,
        label: `📏 ${name}`,
        labelStyle: { fill: "#2dd4bf", fontSize: 10 },
        labelBgStyle: { fill: "#14171c" },
        labelBgPadding: [4, 2] as [number, number],
        zIndex: 20,
        markerEnd: { type: MarkerType.ArrowClosed, color: "#2dd4bf", width: 16, height: 16 },
        style: { stroke: "#2dd4bf", strokeWidth: 2, strokeDasharray: "2 3" },
      });
    });

    // 3) Saved-date references in the answer MUST have a `date` dependency edge
    // (the scheduler derives the serve-by window from it). A valid corpus always
    // has one — but while authoring it's easy to forget. If the date edge is
    // missing, surface the gap as a red warning edge instead of drawing nothing.
    anchorRefsInAnswer(e.answer).forEach((name) => {
      const src = graph.emission_map?.[name];
      if (!src || src === e.id || !graph.emails[src] || !inScenario(src)) return;
      const hasDateEdge = e.depends_on.some((d) => d.email === src && d.type === "date");
      if (hasDateEdge) return; // already drawn as a "deadline" edge above
      noteCrossThread(src, e.id);
      edges.push({
        id: `need-date:${name}:${src}->${e.id}`,
        source: src,
        target: e.id,
        label: `⚠ needs date link · 📌 ${name}`,
        labelStyle: { fill: "#f2647a", fontSize: 10 },
        labelBgStyle: { fill: "#14171c" },
        labelBgPadding: [4, 2] as [number, number],
        zIndex: 21,
        markerEnd: { type: MarkerType.ArrowClosed, color: "#f2647a", width: 16, height: 16 },
        style: { stroke: "#f2647a", strokeWidth: 2, strokeDasharray: "4 3" },
      });
    });
  });

  // 4) Whole-thread (node-level) dependencies: "every email in B comes after
  // every email in A." Drawn once per pair, from A's first email to B's first,
  // and ONLY when no finer value/date/ordering edge already connects the two —
  // otherwise it's just noise on top of a link the author can already see.
  graph.threads.forEach((t) => {
    if (t.scenario !== activeScenario) return;
    (t.node_depends_on || []).forEach((d) => {
      if (!d.node) return;
      const anc = graph.threads.find((x) => x.id === d.node);
      if (!anc || anc.id === t.id || anc.scenario !== activeScenario) return;
      if (linkedThreadPairs.has(pairKey(anc.id, t.id))) return;
      const from = anc.emails[0];
      const to = t.emails[0];
      if (!from || !to || !graph.emails[from] || !graph.emails[to]) return;
      edges.push({
        id: `node:${anc.id}->${t.id}`,
        source: from,
        target: to,
        label: "whole thread: after",
        labelStyle: { fill: "#6b7280", fontSize: 10 },
        labelBgStyle: { fill: "#14171c" },
        labelBgPadding: [4, 2] as [number, number],
        markerEnd: { type: MarkerType.ArrowClosed, color: "#39414e", width: 14, height: 14 },
        style: { stroke: "#39414e", strokeWidth: 1, strokeDasharray: "1 4" },
      });
    });
  });

  return { nodes, edges };
}

const LEGEND_EDGES = [
  { c: "#5a6373", dash: "", label: "after — must happen later in time" },
  { c: "#b98cff", dash: "5 4", label: "deadline — must be done by a date set there" },
  { c: "#2dd4bf", dash: "2 3", label: "📏 shares a value — reuses a number defined there" },
  { c: "#39414e", dash: "1 4", label: "whole thread: after — every email comes later" },
  { c: "#f2647a", dash: "4 3", label: "⚠ needs date link — referenced, but no edge yet" },
];
const LEGEND_TAGS = [
  { t: "📌", label: "sets a saved date others reuse" },
  { t: "📏", label: "defines a shared value others reuse" },
  { t: "N actions", label: "expected actions the secretary must take" },
  { t: "no action", label: "graded as 'correctly do nothing'" },
];

function Legend() {
  const [open, setOpen] = useState(true);
  return (
    <Panel position="top-right">
      <div className="border border-edge bg-panel/95 text-xs text-white/85 shadow-2xl backdrop-blur">
        <button
          onClick={() => setOpen((o) => !o)}
          className="flex w-full items-center gap-2 border-b border-edge px-3 py-1.5 text-[11px] uppercase tracking-wide text-muted hover:text-white"
        >
          <span>{open ? "▾" : "▸"}</span> Legend
        </button>
        {open && (
          <div className="space-y-2 p-3">
            <div className="space-y-1.5">
              {LEGEND_EDGES.map((e) => (
                <div key={e.label} className="flex items-center gap-2">
                  <svg width="34" height="8" className="shrink-0">
                    <line x1="0" y1="4" x2="34" y2="4" stroke={e.c} strokeWidth="2" strokeDasharray={e.dash || undefined} />
                  </svg>
                  <span>{e.label}</span>
                </div>
              ))}
            </div>
            <div className="space-y-1.5 border-t border-edge pt-2">
              {LEGEND_TAGS.map((t) => (
                <div key={t.label} className="flex items-center gap-2">
                  <span className="inline-block w-[58px] shrink-0 border border-edge px-1 py-0.5 text-center text-[10px] text-muted">{t.t}</span>
                  <span>{t.label}</span>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </Panel>
  );
}

export function GraphCanvas() {
  const { graph, selected, select, connect, disconnect, setPosition, oracle, addFollowUp, activeScenario } = useStore();

  const initial = useMemo(
    () => (graph ? build(graph, selected, oracle, addFollowUp, activeScenario) : { nodes: [], edges: [] }),
    // rebuild on structural/content change
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [structureSig(graph), selected, oracle, addFollowUp, activeScenario]
  );

  const [nodes, setNodes, onNodesChange] = useNodesState<RFNode<EmailNodeData>>(initial.nodes);
  const [edges, setEdges, onEdgesChange] = useEdgesState<RFEdge>(initial.edges);

  useEffect(() => {
    setNodes(initial.nodes);
    setEdges(initial.edges);
  }, [initial, setNodes, setEdges]);

  const handleNodesChange = useCallback(
    (changes: NodeChange<RFNode<EmailNodeData>>[]) => {
      onNodesChange(changes);
      changes.forEach((c) => {
        if (c.type === "position" && c.dragging === false && c.position) {
          setPosition(c.id, Math.round(c.position.x), Math.round(c.position.y));
        }
      });
    },
    [onNodesChange, setPosition]
  );

  const onConnect = useCallback((c: Connection) => {
    if (c.source && c.target) connect(c.source, c.target);
  }, [connect]);

  const onEdgesDelete = useCallback((removed: RFEdge[]) => {
    // only explicit ordering edges are author-owned; derived value/date/node
    // edges have no `depends_on` row to remove (id prefix tells them apart).
    removed.forEach((e) => {
      if (e.id.startsWith("dep:")) disconnect(e.target, e.source);
    });
  }, [disconnect]);

  if (!graph) return <div className="flex h-full items-center justify-center text-muted">Loading…</div>;

  return (
    <ReactFlow
      nodes={nodes}
      edges={edges}
      nodeTypes={nodeTypes}
      onNodesChange={handleNodesChange}
      onEdgesChange={onEdgesChange}
      onConnect={onConnect}
      onEdgesDelete={onEdgesDelete}
      onNodeClick={(_, n) => select(n.id)}
      onPaneClick={() => select(null)}
      fitView
      proOptions={{ hideAttribution: true }}
      minZoom={0.2}
    >
      <Background variant={BackgroundVariant.Dots} gap={24} size={1} color="#1b1f26" />
      <Legend />
      <MiniMap
        pannable
        zoomable
        nodeColor={(n) => (n.data as EmailNodeData).color}
        maskColor="#0b0d11cc"
        style={{ background: "#14171c", border: "1px solid #262b34", width: 150, height: 100 }}
      />
      <Controls showInteractive={false} />
    </ReactFlow>
  );
}

// a cheap signature so we only rebuild RF state on meaningful change
function structureSig(graph: Graph | null): string {
  if (!graph) return "";
  const emails = Object.values(graph.emails)
    .map((e) => `${e.id}:${e.thread}:${e.subject}:${JSON.stringify(e.answer)}:${e.depends_on.map((d) => d.email + d.type).join(",")}:${e.body_segments.length}`)
    .join("|");
  const nodeDeps = graph.threads
    .map((t) => `${t.id}>${(t.node_depends_on || []).map((d) => d.node).join(",")}`)
    .join(";");
  return `${emails}::nodes:${nodeDeps}::emit:${Object.keys(graph.emission_map).join(",")}::fact:${Object.keys(graph.fact_map || {}).join(",")}`;
}
