import { useStore } from "../store";
import type { AnswerForm, Segment } from "../types";
import { BodyComposer } from "./BodyComposer";
import { GradingEditor } from "./GradingEditor";

export function EmailEditor() {
  const {
    graph, selected, select, updateEmail, updateAnswer, deleteEmail, disconnect,
    ensureDateEdge, connect, scenarios, activeScenario, moveThreadToScenario, createScenario,
  } = useStore();
  if (!graph || !selected) return null;
  const email = graph.emails[selected];
  if (!email) return null;

  const thread = graph.threads.find((t) => t.id === email.thread);
  const castKeys = Object.keys(thread?.cast || {});
  const serve = graph.sample.serve_date[email.id] || graph.start;
  const anchors = graph.sample.anchors;

  const scenarioOf = (eid: string) => {
    const e = graph.emails[eid];
    const t = e && graph.threads.find((x) => x.id === e.thread);
    return t?.scenario ?? "";
  };
  const myScenario = thread?.scenario ?? activeScenario;
  // candidates to depend ON: any other email not already a dependency.
  const depCandidates = Object.values(graph.emails)
    .filter((e) => e.id !== email.id && !email.depends_on.some((d) => d.email === e.id))
    .sort((a, b) => scenarioOf(a.id).localeCompare(scenarioOf(b.id)) || a.id.localeCompare(b.id));

  const moveValue = "__current";

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
        <Labeled label="Scenario (tab)" className="mb-3">
          <select
            className={inp}
            value={moveValue}
            onChange={(e) => {
              const v = e.target.value;
              if (!thread) return;
              if (v === "__new") {
                const name = createScenario();
                moveThreadToScenario(thread.id, name);
              } else if (v !== moveValue) {
                moveThreadToScenario(thread.id, v);
              }
            }}
          >
            <option value={moveValue}>{myScenario === "unsorted" ? "Unsorted" : myScenario} (current)</option>
            {scenarios.filter((s) => s !== myScenario).map((s) => (
              <option key={s} value={s}>move thread → {s === "unsorted" ? "Unsorted" : s}</option>
            ))}
            <option value="__new">move thread → new scenario…</option>
          </select>
        </Labeled>

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

        <div className="mt-3">
          <div className="mb-1 text-[11px] uppercase tracking-wide text-muted">Depends on</div>
          {email.depends_on.length > 0 && (
            <div className="mb-1.5 flex flex-wrap gap-1.5">
              {email.depends_on.map((d) => (
                <span key={d.email} className={`inline-flex items-center gap-1 rounded-md border px-2 py-0.5 text-xs ${
                  d.type === "date" ? "border-anchor/50 bg-anchor/10 text-anchor" : "border-edge bg-panel2 text-muted"}`}>
                  {d.type === "date" ? "⏰" : "🔗"} {d.email}
                  <button onClick={() => disconnect(email.id, d.email!)} className="hover:text-bad">×</button>
                </span>
              ))}
            </div>
          )}
          <select
            className={inp}
            value=""
            onChange={(e) => { if (e.target.value) connect(e.target.value, email.id); }}
          >
            <option value="">+ depends on…</option>
            {depCandidates.map((c) => {
              const cross = scenarioOf(c.id) !== myScenario;
              return (
                <option key={c.id} value={c.id}>
                  {c.id}{cross ? `  (merges ${scenarioOf(c.id) === "unsorted" ? "Unsorted" : scenarioOf(c.id)})` : ""}
                </option>
              );
            })}
          </select>
          <div className="mt-1 text-[10px] text-muted">Picking an email from another tab merges the two scenarios.</div>
        </div>

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
