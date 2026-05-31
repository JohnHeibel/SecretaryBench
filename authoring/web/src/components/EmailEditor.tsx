import { useStore } from "../store";
import type { AnswerForm, Segment } from "../types";
import { BodyComposer } from "./BodyComposer";
import { GradingEditor } from "./GradingEditor";

export function EmailEditor() {
  const { graph, selected, select, updateEmail, updateAnswer, deleteEmail, disconnect, ensureDateEdge } = useStore();
  if (!graph || !selected) return null;
  const email = graph.emails[selected];
  if (!email) return null;

  const thread = graph.threads.find((t) => t.id === email.thread);
  const castKeys = Object.keys(thread?.cast || {});
  const serve = graph.sample.serve_date[email.id] || graph.start;
  const anchors = graph.sample.anchors;

  return (
    <div className="flex h-full w-[560px] shrink-0 flex-col border-l border-edge bg-panel scroll-thin">
      <div className="flex items-center justify-between border-b border-edge px-4 py-3">
        <div>
          <div className="font-mono text-xs text-muted">{email.thread}</div>
          <div className="font-mono text-sm text-white">{email.id}</div>
        </div>
        <div className="flex gap-2">
          <button onClick={() => deleteEmail(email.id)} className="rounded-md border border-edge px-2 py-1 text-xs text-muted hover:border-bad hover:text-bad">
            delete email
          </button>
          <button onClick={() => select(null)} className="rounded-md border border-edge px-2 py-1 text-xs text-muted hover:text-white">
            close
          </button>
        </div>
      </div>

      <div className="flex-1 overflow-y-auto scroll-thin p-4">
        <div className="grid grid-cols-2 gap-3">
          <Labeled label="From">
            <input list="cast" className={inp} value={email.from} onChange={(e) => updateEmail(email.id, { from: e.target.value })} />
          </Labeled>
          <Labeled label="To">
            <input list="cast" className={inp} value={email.to.join(", ")}
              onChange={(e) => updateEmail(email.id, { to: e.target.value.split(",").map((s) => s.trim()).filter(Boolean) })} />
          </Labeled>
        </div>
        <datalist id="cast">{castKeys.map((k) => <option key={k} value={k} />)}</datalist>

        <Labeled label="Subject" className="mt-3">
          <input className={inp} value={email.subject} onChange={(e) => updateEmail(email.id, { subject: e.target.value })} />
        </Labeled>

        {email.depends_on.length > 0 && (
          <div className="mt-3">
            <div className="mb-1 text-[11px] uppercase tracking-wide text-muted">Depends on</div>
            <div className="flex flex-wrap gap-1.5">
              {email.depends_on.map((d) => (
                <span key={d.email} className={`inline-flex items-center gap-1 rounded-md border px-2 py-0.5 text-xs ${
                  d.type === "date" ? "border-anchor/50 bg-anchor/10 text-anchor" : "border-edge bg-panel2 text-muted"}`}>
                  {d.type === "date" ? "⏰" : "🔗"} {d.email}
                  <button onClick={() => disconnect(email.id, d.email!)} className="hover:text-bad">×</button>
                </span>
              ))}
            </div>
          </div>
        )}

        <div className="mt-4">
          <div className="mb-1 text-[11px] uppercase tracking-wide text-muted">Email body</div>
          <BodyComposer email={email} serve={serve} anchors={anchors}
            onChange={(segments: Segment[]) => updateEmail(email.id, { body_segments: segments })} />
        </div>

        <div className="mt-5 border-t border-edge pt-4">
          <GradingEditor email={email} serve={serve} anchors={anchors}
            onChange={(a: AnswerForm) => updateAnswer(email.id, a)}
            onAnchorPicked={(name) => ensureDateEdge(email.id, name)} />
        </div>
      </div>
    </div>
  );
}

function Labeled({ label, children, className }: { label: string; children: React.ReactNode; className?: string }) {
  return (
    <label className={`block ${className || ""}`}>
      <div className="mb-1 text-[11px] uppercase tracking-wide text-muted">{label}</div>
      {children}
    </label>
  );
}

const inp = "w-full rounded-md border border-edge bg-ink px-2 py-1 text-sm text-white outline-none focus:border-accent";
