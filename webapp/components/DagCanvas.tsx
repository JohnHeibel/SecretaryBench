"use client";
import { useMemo } from "react";
import { ReactFlow, Background, Controls, MarkerType, type Edge as RFEdge, type Node as RFNode } from "@xyflow/react";
import type { CorpusNode, LintResult } from "@/lib/types";

interface Props {
  nodes: CorpusNode[];
  lint: LintResult | null;
  serveDate: string;
  onSelectEmail: (nodeId: string, emailId: string) => void;
}

// longest-path depth per email, so prerequisites sit to the left of dependents
function layer(emails: { id: string; deps: string[] }[]): Record<string, number> {
  const byId = new Map(emails.map((e) => [e.id, e]));
  const depth: Record<string, number> = {};
  const visiting = new Set<string>();
  const calc = (id: string): number => {
    if (depth[id] !== undefined) return depth[id];
    if (visiting.has(id)) return 0; // cycle guard; lint will flag it
    visiting.add(id);
    const e = byId.get(id);
    const d = !e || e.deps.length === 0 ? 0 : 1 + Math.max(...e.deps.map((p) => (byId.has(p) ? calc(p) : -1)));
    visiting.delete(id);
    return (depth[id] = Math.max(0, d));
  };
  for (const e of emails) calc(e.id);
  return depth;
}

export default function DagCanvas({ nodes, lint, serveDate, onSelectEmail }: Props) {
  const palette = ["#0ea5e9", "#a855f7", "#22c55e", "#f59e0b", "#ef4444", "#14b8a6", "#ec4899"];
  const nodeColor = useMemo(() => {
    const m: Record<string, string> = {};
    nodes.forEach((n, i) => (m[n.id] = palette[i % palette.length]));
    return m;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [nodes]);

  const { rfNodes, rfEdges } = useMemo(() => {
    const flat = nodes.flatMap((n) => n.emails.map((e) => ({ ...e, _node: n.id })));
    const emitByEmail: Record<string, string[]> = {};
    for (const le of lint?.emails ?? []) emitByEmail[le.id] = le.emits;

    const depth = layer(flat.map((e) => ({ id: e.id, deps: e.depends_on.map((d) => d.email).filter(Boolean) as string[] })));
    const perCol: Record<number, number> = {};

    const rfNodes: RFNode[] = flat.map((e) => {
      const col = depth[e.id] ?? 0;
      const row = (perCol[col] = (perCol[col] ?? 0) + 1) - 1;
      const emits = emitByEmail[e.id] ?? [];
      return {
        id: e.id,
        position: { x: col * 260, y: row * 104 },
        data: {
          label: (
            <div className="text-left">
              <div className="font-mono text-[10px] text-slate-400">{e.id}</div>
              <div className="text-xs font-medium text-slate-100">{e.subject || "(no subject)"}</div>
              {emits.length > 0 && <div className="mt-0.5 text-[10px] text-violet-300">emits @{emits.join(", @")}</div>}
              {(e.answer?.expect?.length ?? 0) === 0 && <div className="text-[10px] text-slate-500">no-action</div>}
            </div>
          ),
        },
        style: { background: "#0f172a", border: `1px solid ${nodeColor[e._node]}`, borderRadius: 10, padding: 8, width: 220 },
      };
    });

    const rfEdges: RFEdge[] = flat.flatMap((e) =>
      e.depends_on.filter((d) => d.email).map((d) => ({
        id: `${d.email}->${e.id}`,
        source: d.email!,
        target: e.id,
        animated: d.type === "date",
        label: d.type,
        labelStyle: { fontSize: 9, fill: d.type === "date" ? "#38bdf8" : "#f59e0b" },
        style: { stroke: d.type === "date" ? "#38bdf8" : "#f59e0b", strokeWidth: 1.5 },
        markerEnd: { type: MarkerType.ArrowClosed, color: d.type === "date" ? "#38bdf8" : "#f59e0b" },
      })),
    );
    return { rfNodes, rfEdges };
  }, [nodes, lint, nodeColor]);

  const emailNode = (id: string) => {
    for (const n of nodes) if (n.emails.some((e) => e.id === id)) return n.id;
    return null;
  };

  return (
    <div className="h-full w-full">
      <div className="flex items-center gap-4 border-b border-slate-800 bg-slate-900/60 px-4 py-1.5 text-[11px] text-slate-400">
        <span>serve preview {serveDate}</span>
        <span className="flex items-center gap-1"><span className="inline-block h-0.5 w-4" style={{ background: "#f59e0b" }} /> static (retrieval span)</span>
        <span className="flex items-center gap-1"><span className="inline-block h-0.5 w-4" style={{ background: "#38bdf8" }} /> date (deadline)</span>
        <span className="text-violet-300">@anchor = published date</span>
      </div>
      <div style={{ height: "calc(100% - 30px)" }}>
        <ReactFlow nodes={rfNodes} edges={rfEdges} fitView minZoom={0.2}
          onNodeClick={(_, n) => { const nid = emailNode(n.id); if (nid) onSelectEmail(nid, n.id); }}>
          <Background color="#1e293b" gap={20} />
          <Controls />
        </ReactFlow>
      </div>
    </div>
  );
}
