import {
  Background,
  BackgroundVariant,
  Controls,
  MiniMap,
  ReactFlow,
  useEdgesState,
  useNodesState,
  type Connection,
  type Edge as RFEdge,
  type Node as RFNode,
  type NodeChange,
} from "@xyflow/react";
import "@xyflow/react/dist/style.css";
import { useCallback, useEffect, useMemo } from "react";
import { chipLabel } from "../chipLabel";
import { useStore } from "../store";
import type { Graph } from "../types";
import { EmailNode, type EmailNodeData } from "./EmailNode";

const THREAD_COLORS = ["#6aa6ff", "#34d399", "#fbbf24", "#fb7185", "#c084fc", "#22d3ee", "#f472b6", "#a3e635"];
const nodeTypes = { email: EmailNode };

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
  onFollowUp: (id: string) => void
) {
  const colorOf: Record<string, string> = {};
  graph.threads.forEach((t, i) => (colorOf[t.id] = THREAD_COLORS[i % THREAD_COLORS.length]));
  const emitsBy: Record<string, string[]> = {};
  Object.entries(graph.emission_map).forEach(([name, eid]) => {
    (emitsBy[eid] ||= []).push(name);
  });
  const factsBy: Record<string, string[]> = {};
  Object.entries(graph.fact_map || {}).forEach(([name, eid]) => {
    (factsBy[eid] ||= []).push(name);
  });

  const nodes: RFNode<EmailNodeData>[] = Object.values(graph.emails).map((e) => ({
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
  Object.values(graph.emails).forEach((e) => {
    e.depends_on.forEach((d) => {
      if (!d.email || !graph.emails[d.email]) return;
      const isDate = d.type === "date";
      edges.push({
        id: `${d.email}->${e.id}`,
        source: d.email,
        target: e.id,
        animated: isDate,
        label: isDate ? "deadline" : "after",
        labelStyle: { fill: isDate ? "#b98cff" : "#888f9c", fontSize: 10 },
        labelBgStyle: { fill: "#14171c" },
        labelBgPadding: [4, 2] as [number, number],
        style: { stroke: isDate ? "#b98cff" : "#39414e", strokeDasharray: isDate ? "5 4" : undefined },
      });
    });
  });
  return { nodes, edges };
}

export function GraphCanvas() {
  const { graph, selected, select, connect, disconnect, setPosition, oracle, addFollowUp } = useStore();

  const initial = useMemo(
    () => (graph ? build(graph, selected, oracle, addFollowUp) : { nodes: [], edges: [] }),
    // rebuild on structural/content change
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [structureSig(graph), selected, oracle, addFollowUp]
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
    removed.forEach((e) => disconnect(e.target, e.source));
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
  return Object.values(graph.emails)
    .map((e) => `${e.id}:${e.thread}:${e.subject}:${e.answer.expect.length}:${e.depends_on.map((d) => d.email + d.type).join(",")}:${e.body_segments.length}`)
    .join("|") + "::" + Object.keys(graph.emission_map).join(",");
}
