import { Handle, Position } from "@xyflow/react";
import type { EmailForm } from "../types";

export interface EmailNodeData {
  email: EmailForm;
  color: string;
  selected: boolean;
  oracle?: { passed: boolean; headline: string };
  emitsAnchors: string[];
  definesFacts: string[];
  bodyPreview: string;
  onFollowUp: (id: string) => void;
  [key: string]: unknown;
}

export function EmailNode({ data }: { data: EmailNodeData }) {
  const { email, color, oracle, emitsAnchors, definesFacts, bodyPreview } = data;
  const actions = email.answer.expect.length;
  const noAction = actions === 0 && email.answer.forbid.length === 0;

  return (
    <div className={`group relative w-[252px] border bg-panel ${data.selected ? "border-accent" : "border-edge"}`}>
      <Handle type="target" position={Position.Top} />

      <div className="flex items-center gap-2 px-3 py-1.5" style={{ borderLeft: `3px solid ${color}` }}>
        <span className="truncate font-mono text-[11px] text-muted">{email.id}</span>
        {oracle && (
          <span className={`ml-auto text-[10px] ${oracle.passed ? "text-ok" : "text-bad"}`} title={oracle.headline}>
            {oracle.passed ? "solvable" : "unsolvable"}
          </span>
        )}
      </div>

      <div className="border-t border-edge px-3 py-2">
        <div className="truncate text-[13px] font-semibold text-white">
          {email.subject || <span className="font-normal text-muted">(no subject)</span>}
        </div>
        <div className="mt-0.5 line-clamp-2 text-xs leading-snug text-muted">{bodyPreview || "(empty)"}</div>
        <div className="mt-2 flex flex-wrap items-center gap-1">
          {noAction
            ? <Tag className="border-edge text-muted">no action</Tag>
            : <Tag className="border-accent/40 text-accent">{actions} action{actions === 1 ? "" : "s"}</Tag>}
          {emitsAnchors.map((a) => <Tag key={a} className="border-anchor/50 text-anchor">📌 {a}</Tag>)}
          {definesFacts.map((f) => <Tag key={f} className="border-fact/50 text-fact">📏 {f}</Tag>)}
        </div>
      </div>

      <Handle type="source" position={Position.Bottom} />

      {/* obvious way to extend a conversation */}
      <button
        className="nodrag absolute -bottom-3 left-1/2 z-10 -translate-x-1/2 whitespace-nowrap border border-edge bg-panel2 px-2 py-0.5 text-[11px] text-muted opacity-0 transition group-hover:opacity-100 hover:border-accent hover:text-accent"
        onClick={(e) => { e.stopPropagation(); data.onFollowUp(email.id); }}
        title="Add a reply / follow-up that depends on this email"
      >
        + follow-up
      </button>
    </div>
  );
}

function Tag({ children, className }: { children: React.ReactNode; className?: string }) {
  return <span className={`border px-1.5 py-0.5 text-[10px] ${className || ""}`}>{children}</span>;
}
