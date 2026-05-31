import { Handle, Position } from "@xyflow/react";
import type { EmailForm } from "../types";

export interface EmailNodeData {
  email: EmailForm;
  color: string;
  selected: boolean;
  oracle?: { passed: boolean; headline: string };
  emitsAnchors: string[];
  bodyPreview: string;
  [key: string]: unknown;
}

export function EmailNode({ data }: { data: EmailNodeData }) {
  const { email, color, oracle, emitsAnchors, bodyPreview } = data;
  const actions = email.answer.expect.length;
  const noAction = actions === 0 && email.answer.forbid.length === 0;

  return (
    <div
      className={`w-[260px] rounded-xl border-2 bg-panel shadow-lg transition ${
        data.selected ? "border-accent ring-2 ring-accent/40" : "border-edge"
      }`}
    >
      <Handle type="target" position={Position.Top} />
      <div className="flex items-center gap-2 rounded-t-[10px] px-3 py-1.5" style={{ background: color + "22" }}>
        <span className="h-2.5 w-2.5 rounded-full" style={{ background: color }} />
        <span className="truncate font-mono text-[11px] text-muted">{email.id}</span>
        {oracle && (
          <span className={`ml-auto text-[11px] ${oracle.passed ? "text-ok" : "text-bad"}`} title={oracle.headline}>
            {oracle.passed ? "● solvable" : "● unsolvable"}
          </span>
        )}
      </div>
      <div className="px-3 py-2">
        <div className="truncate text-sm font-semibold text-white">{email.subject || <span className="text-muted">(no subject)</span>}</div>
        <div className="mt-0.5 line-clamp-2 text-xs text-muted">{bodyPreview || "(empty)"}</div>
        <div className="mt-2 flex flex-wrap items-center gap-1">
          {noAction ? (
            <Tag className="border-edge text-muted">no action</Tag>
          ) : (
            <Tag className="border-accent/40 text-accent">{actions} action{actions === 1 ? "" : "s"}</Tag>
          )}
          {emitsAnchors.map((a) => (
            <Tag key={a} className="border-anchor/50 text-anchor">📌 {a}</Tag>
          ))}
        </div>
      </div>
      <Handle type="source" position={Position.Bottom} />
    </div>
  );
}

function Tag({ children, className }: { children: React.ReactNode; className?: string }) {
  return <span className={`rounded border px-1.5 py-0.5 text-[10px] ${className || ""}`}>{children}</span>;
}
