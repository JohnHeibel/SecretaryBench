"use client";
import { useCallback, useMemo } from "react";
import { ReactFlow, Background, Controls, Handle, Position, MarkerType, type Connection, type Edge as RFEdge, type Node as RFNode, type NodeProps } from "@xyflow/react";
import Dagre from "@dagrejs/dagre";
import type { CorpusNode, LintResult } from "@/lib/types";

interface Props {
  nodes: CorpusNode[];
  lint: LintResult | null;
  serveDate: string;
  onSelectEmail: (nodeId: string, emailId: string) => void;
  // Drag-connect: the author dragged from `prerequisiteId`'s bottom handle to `targetId`'s top handle,
  // so `targetId` should depend_on `prerequisiteId` (the same direction the editor's picker uses). Wired
  // through Workspace -> updateEmail -> withDerivedDateEdges so validation/auto-wiring stays identical.
  onAddEdge?: (nodeId: string, targetId: string, prerequisiteId: string) => void;
}

const NODE_W = 220;
const NODE_H = 104;       // approx card height handed to dagre so ranks don't overlap
const LANE_PAD = 28;      // breathing room around a storyline lane
const LANE_HEADER = 28;   // height reserved for the lane's title strip
const LANE_GAP = 70;      // gap between adjacent storyline lanes

// Top-to-bottom hierarchical layout for ONE storyline via dagre: it assigns ranks (prerequisites above
// dependents), ORDERS each rank to minimize edge crossings, and CENTERS a parent over its children — the
// three things the old hand-rolled "one wide row, parent pinned far left" layout got wrong. Returns each
// email's top-left position (dagre reports centers) plus the storyline's bounding box for lane sizing.
function layoutStoryline(emails: { id: string }[], edges: [string, string][]): { pos: Record<string, { x: number; y: number }>; w: number; h: number } {
  const g = new Dagre.graphlib.Graph();
  g.setGraph({ rankdir: "TB", nodesep: 36, ranksep: 64, marginx: 0, marginy: 0 });
  g.setDefaultEdgeLabel(() => ({}));
  for (const e of emails) g.setNode(e.id, { width: NODE_W, height: NODE_H });
  for (const [src, tgt] of edges) if (g.hasNode(src) && g.hasNode(tgt)) g.setEdge(src, tgt);
  Dagre.layout(g);
  const pos: Record<string, { x: number; y: number }> = {};
  let w = 0, h = 0;
  for (const e of emails) {
    const n = g.node(e.id);
    const x = n.x - NODE_W / 2, y = n.y - NODE_H / 2;   // center -> top-left
    pos[e.id] = { x, y };
    w = Math.max(w, x + NODE_W); h = Math.max(h, y + NODE_H);
  }
  return { pos, w, h };
}

// One email card. Custom node so it carries a TOP target handle and a BOTTOM source handle: dragging
// from a card's bottom down onto another card's top creates the dependency (downward = "comes after").
type EmailData = {
  emailId: string;
  subject: string;
  emits: string[];
  noAction: boolean;
  isReply: boolean;
  followsFrom?: string; // the prerequisite this reply answers (for the "Re:" subtitle)
  accent: string;
};

function EmailCard({ data, selected }: NodeProps) {
  const d = data as EmailData;
  return (
    <div className="relative rounded-[10px] border bg-slate-900 px-2 py-2 text-left shadow-sm"
      style={{ width: NODE_W, borderColor: d.accent, borderLeftWidth: d.isReply ? 4 : 1, outline: selected ? `2px solid ${d.accent}` : "none" }}>
      <Handle type="target" position={Position.Top} className="!h-2 !w-2 !border-slate-500 !bg-slate-700" />
      <div className="flex items-center gap-1.5">
        {d.isReply && <span className="shrink-0 rounded bg-slate-700 px-1 py-px text-[9px] font-semibold uppercase tracking-wide text-slate-300" title="A reply: this email follows from an earlier one in the storyline.">Re</span>}
        <span className="truncate font-mono text-[10px] text-slate-400">{d.emailId}</span>
      </div>
      <div className="mt-0.5 text-xs font-medium text-slate-100">{d.subject || "(no subject)"}</div>
      {d.isReply && d.followsFrom && <div className="mt-0.5 truncate text-[10px] text-slate-500" title={`replies to ${d.followsFrom}`}>↩ {d.followsFrom}</div>}
      {d.emits.length > 0 && <div className="mt-0.5 text-[10px] text-violet-300">emits @{d.emits.join(", @")}</div>}
      {d.noAction && <div className="text-[10px] text-slate-500">no-action</div>}
      <Handle type="source" position={Position.Bottom} className="!h-2 !w-2 !border-slate-500 !bg-slate-700" />
    </div>
  );
}

// A passive storyline lane drawn BEHIND its emails (only when more than one storyline is shown). Just a
// labeled, tinted container so the grouping reads at a glance; it isn't selectable or draggable.
function LaneCard({ data }: NodeProps) {
  const d = data as { title: string; accent: string; width: number; height: number };
  return (
    <div className="rounded-xl border" style={{ width: d.width, height: d.height, borderColor: `${d.accent}55`, background: `${d.accent}10` }}>
      <div className="px-3 py-1 text-[11px] font-semibold tracking-wide" style={{ color: d.accent }}>{d.title}</div>
    </div>
  );
}

const nodeTypes = { email: EmailCard, lane: LaneCard };

export default function DagCanvas({ nodes, lint, serveDate, onSelectEmail, onAddEdge }: Props) {
  const palette = ["#0ea5e9", "#a855f7", "#22c55e", "#f59e0b", "#ef4444", "#14b8a6", "#ec4899"];
  const accentOf = useMemo(() => {
    const m: Record<string, string> = {};
    nodes.forEach((n, i) => (m[n.id] = palette[i % palette.length]));
    return m;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [nodes]);

  const emitByEmail = useMemo(() => {
    const m: Record<string, string[]> = {};
    for (const le of lint?.emails ?? []) m[le.id] = le.emits;
    return m;
  }, [lint]);

  const { rfNodes, rfEdges } = useMemo(() => {
    const rfNodes: RFNode[] = [];
    const rfEdges: RFEdge[] = [];
    const multi = nodes.length > 1; // only draw storyline lanes when there's more than one to tell apart
    let laneX = 0; // lanes are laid side by side; emails are positioned by dagre within their lane

    for (const node of nodes) {
      const accent = accentOf[node.id];
      const ids = new Set(node.emails.map((e) => e.id));
      // An email is a REPLY if its subject starts with "Re:" (case-insensitive) OR it depends on an
      // earlier email in this same storyline. followsFrom names the in-storyline prerequisite it answers.
      const isReply = (e: { subject: string; depends_on: { email?: string }[] }) =>
        /^\s*re:/i.test(e.subject) || e.depends_on.some((d) => d.email && ids.has(d.email));
      const followsFrom = (e: { depends_on: { email?: string }[] }) => e.depends_on.map((d) => d.email).find((id) => id && ids.has(id)) ?? undefined;

      // Every in-storyline depends_on edge (static OR date) defines "comes after", so all of them feed the
      // dagre ranking; that's what places prerequisites above dependents and centers them.
      const intra: [string, string][] = [];
      for (const e of node.emails) for (const dep of e.depends_on) if (dep.email && ids.has(dep.email)) intra.push([dep.email, e.id]);
      const { pos, w, h } = layoutStoryline(node.emails, intra);

      if (multi) {
        rfNodes.push({ id: `lane:${node.id}`, type: "lane", position: { x: laneX, y: 0 }, selectable: false, draggable: false, connectable: false, zIndex: -1,
          data: { title: node.id, accent, width: w + LANE_PAD * 2, height: h + LANE_HEADER + LANE_PAD } });
      }
      const ox = multi ? laneX + LANE_PAD : 0;
      const oy = multi ? LANE_HEADER : 0;

      for (const e of node.emails) {
        const reply = isReply(e);
        rfNodes.push({
          id: e.id, type: "email", position: { x: ox + pos[e.id].x, y: oy + pos[e.id].y },
          data: { emailId: e.id, subject: e.subject, emits: emitByEmail[e.id] ?? [], noAction: (e.answer?.ops?.length ?? 0) === 0, isReply: reply, followsFrom: reply ? followsFrom(e) : undefined, accent } satisfies EmailData,
        });
        for (const dep of e.depends_on) {
          if (!dep.email) continue;
          const isDate = dep.type === "date";
          // Date edges (needles) are the many "reuses an earlier date" links; draw them thin and a bit
          // translucent so a hub email feeding a dozen needles doesn't become a wall of lines. Static
          // ("comes after" / reply) edges define the thread shape, so they stay solid and opaque.
          const color = isDate ? "#38bdf8" : "#f59e0b";
          rfEdges.push({
            id: `${dep.email}->${e.id}`, source: dep.email, target: e.id, type: "default", animated: false,
            style: { stroke: color, strokeWidth: isDate ? 1 : 1.75, strokeDasharray: isDate ? "4 4" : undefined, opacity: isDate ? 0.45 : 0.9 },
            markerEnd: { type: MarkerType.ArrowClosed, color, width: 14, height: 14 },
          });
        }
      }
      laneX += w + LANE_PAD * 2 + LANE_GAP; // start of the next storyline lane
    }
    return { rfNodes, rfEdges };
  }, [nodes, accentOf, emitByEmail]);

  const emailNode = useCallback((id: string) => {
    for (const n of nodes) if (n.emails.some((e) => e.id === id)) return n.id;
    return null;
  }, [nodes]);

  // Drag-connect: `source` (bottom handle, the prerequisite) -> `target` (top handle, the dependent).
  // Only connect within one storyline (cross-node edges are an advanced case the editor doesn't offer).
  const onConnect = useCallback((c: Connection) => {
    if (!onAddEdge || !c.source || !c.target || c.source === c.target) return;
    const tnode = emailNode(c.target);
    if (tnode && tnode === emailNode(c.source)) onAddEdge(tnode, c.target, c.source);
  }, [onAddEdge, emailNode]);

  return (
    <div className="h-full w-full">
      <div className="flex flex-wrap items-center gap-x-4 gap-y-1 border-b border-slate-800 bg-slate-900/60 px-4 py-1.5 text-[11px] text-slate-400">
        <span>preview date {serveDate}</span>
        <span className="flex items-center gap-1"><svg width="20" height="6"><line x1="0" y1="3" x2="20" y2="3" stroke="#f59e0b" strokeWidth="1.75" /></svg> comes after</span>
        <span className="flex items-center gap-1"><svg width="20" height="6"><line x1="0" y1="3" x2="20" y2="3" stroke="#38bdf8" strokeWidth="1" strokeDasharray="4 4" /></svg> reuses an earlier date (needle)</span>
        <span className="text-violet-300">@anchor = a published date</span>
        <span className="flex items-center gap-1"><span className="rounded bg-slate-700 px-1 py-px text-[9px] font-semibold uppercase text-slate-300">Re</span> a reply</span>
        {onAddEdge && <span className="text-slate-500">drag a card's bottom dot onto another card to link them</span>}
      </div>
      <div style={{ height: "calc(100% - 30px)" }}>
        <ReactFlow nodes={rfNodes} edges={rfEdges} nodeTypes={nodeTypes} fitView minZoom={0.15}
          onConnect={onConnect}
          onNodeClick={(_, n) => { if (n.type !== "email") return; const nid = emailNode(n.id); if (nid) onSelectEmail(nid, n.id); }}>
          <Background color="#1e293b" gap={20} />
          <Controls />
        </ReactFlow>
      </div>
    </div>
  );
}
