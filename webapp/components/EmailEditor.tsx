"use client";
import { useEffect, useState } from "react";
import type { CorpusNode, Edge, EdgeType, Email } from "@/lib/types";
import { anchorOrigins, bodyAnchorNames, obligationKinds, obligationNames } from "@/lib/grammar";
import { STANDARD_ROSTER, normalizeKey, normalizeName } from "@/lib/people";
import BodyEditor from "./BodyEditor";
import AnswerKeyBuilder from "./AnswerKeyBuilder";

interface Props {
  node: CorpusNode;
  email: Email | null;
  allNodes: CorpusNode[];
  anchors: string[];
  serveDate: string;
  onUpdateNode: (node: CorpusNode) => void;
  onRenameNode: (oldId: string, newId: string) => boolean;
  onAddEmail: () => void;
  onUpdateEmail: (email: Email) => void;
  onAutoSlugEmail: (email: Email) => void;
}

export default function EmailEditor({ node, email, allNodes, anchors, serveDate, onUpdateNode, onRenameNode, onAddEmail, onUpdateEmail, onAutoSlugEmail }: Props) {
  return (
    <div className="mx-auto max-w-3xl space-y-5 p-5">
      <Primer />
      <label className="flex items-center gap-2 text-xs font-semibold uppercase tracking-wide text-slate-400"
        title="A storyline is a group of related emails (one scenario). This is its name; rename it freely.">
        Storyline name
        <input key={node.id} defaultValue={node.id} spellCheck={false}
          onKeyDown={(e) => { if (e.key === "Enter") (e.target as HTMLInputElement).blur(); }}
          onBlur={(e) => { if (!onRenameNode(node.id, e.target.value)) e.target.value = node.id; }}
          className="flex-1 rounded border border-slate-700 bg-slate-800 px-2 py-1 font-mono text-sm normal-case tracking-normal text-slate-100" />
      </label>
      {email ? (
        <EmailPanel node={node} email={email} allNodes={allNodes} anchors={anchors} serveDate={serveDate} onUpdateNode={onUpdateNode} onUpdateEmail={onUpdateEmail} onAutoSlugEmail={onAutoSlugEmail} />
      ) : (
        <FirstEmailEmpty nodeId={node.id} onAddEmail={onAddEmail} />
      )}
    </div>
  );
}

function FirstEmailEmpty({ nodeId, onAddEmail }: { nodeId: string; onAddEmail: () => void }) {
  return (
    <div className="rounded-lg border border-dashed border-slate-700 bg-slate-900/35 px-5 py-8 text-center">
      <p className="text-xs font-semibold uppercase tracking-wide text-sky-300">Next step</p>
      <h2 className="mt-2 text-lg font-semibold text-slate-100">Add the first email to {nodeId}.</h2>
      <p className="mx-auto mt-2 max-w-md text-sm leading-6 text-slate-400">Start with a real email someone might send. After you write it, the answer key says what a perfect assistant should do.</p>
      <button onClick={onAddEmail} className="mt-5 rounded-md bg-sky-600 px-4 py-2 text-sm font-semibold text-white hover:bg-sky-500">Add first email</button>
    </div>
  );
}

// A one-read orientation for first-time authors. Collapses and remembers the choice,
// so returning authors aren't nagged. The deep version is the in-app /guide page.
function Primer() {
  const [open, setOpen] = useState(true);
  useEffect(() => { setOpen(localStorage.getItem("sb_primer_closed") !== "1"); }, []);
  function close() { setOpen(false); localStorage.setItem("sb_primer_closed", "1"); }
  if (!open) return (
    <button onClick={() => setOpen(true)} className="text-xs text-slate-500 hover:text-sky-300">▸ How this works</button>
  );
  return (
    <div className="rounded-lg border border-sky-900/60 bg-sky-950/30 p-3 text-xs leading-relaxed text-slate-300">
      <div className="mb-1 flex items-center justify-between">
        <span className="font-semibold text-sky-300">How this works</span>
        <button onClick={close} className="text-slate-500 hover:text-slate-300">hide ✕</button>
      </div>
      <p>You&apos;re writing an email to a busy exec&apos;s <strong>AI assistant</strong>, then saying what a <em>perfect</em> assistant should do with it. Some emails need an action, like putting something on the calendar, adding a to-do, or moving or canceling one. Others need <strong>nothing</strong> done at all.</p>
      <p className="mt-1 text-slate-400">You&apos;re inside a <strong className="text-slate-300">storyline</strong>, a group of related emails. Each email is <strong className="text-slate-300">from</strong> someone — a colleague, a client, or you (the CEO) — and it&apos;s always written <strong className="text-slate-300">to you</strong>, the boss the assistant works for.</p>
      <ol className="ml-4 mt-1 list-decimal space-y-0.5">
        <li>Write the email (who it&apos;s from, the subject, the body). Use <span className="font-mono text-sky-300">+ insert date</span> for any date so it stays exact.</li>
        <li>In <strong>Answer key</strong>, say what to do: name the thing, pick its date, or tick <em>&ldquo;this email needs no action.&rdquo;</em></li>
        <li>The bar at the bottom should read <span className="text-emerald-400">Ready for export</span> and <span className="text-emerald-400">oracle solves 100%</span>. That means a perfect assistant could actually do it.</li>
      </ol>
      <p className="mt-1 text-slate-400">Want more? <a href="/guide" className="text-sky-400 underline hover:text-sky-300">Open the full walkthrough →</a> (worked examples, the date builder, how to build a needle)</p>
    </div>
  );
}

function EmailPanel({ node, email, allNodes, anchors, serveDate, onUpdateNode, onUpdateEmail, onAutoSlugEmail }: Omit<Props, "onRenameNode" | "onAddEmail"> & { email: Email }) {
  const set = (p: Partial<Email>) => onUpdateEmail({ ...email, ...p });
  // Where each published date came from (provenance labels), every date written in any email body
  // (one-click "reuse" chips — same-email reuse is valid and the common case), and this email's OWN
  // body anchors (the "scaffold an action per date" source).
  const origins = anchorOrigins(allNodes);
  const ownBodyAnchors = bodyAnchorNames(email.body);

  return (
    <div className="space-y-4">
      <section className="space-y-3">
        <StepTitle n={1} title="Write the email" note="Use normal email prose. Add any date with the date builder." />
      <div className="flex flex-wrap items-center gap-2 text-xs text-slate-500">
        <span className="truncate font-medium text-slate-300">{email.subject || "(no subject yet, add one below)"}</span>
        <span className="shrink-0 rounded bg-slate-800 px-1.5 py-0.5 font-mono text-[10px] text-slate-500" title="id for this email, made from the subject and used by dependency links">{email.id}</span>
      </div>

      <SenderPicker node={node} email={email} onUpdateNode={onUpdateNode} />

      <label className="block text-xs text-slate-400">Subject
        <input value={email.subject} onChange={(e) => set({ subject: e.target.value })} onBlur={() => onAutoSlugEmail(email)}
          className="mt-1 w-full rounded border border-slate-700 bg-slate-800 px-3 py-1.5 text-sm text-slate-100" />
      </label>

      <BodyEditor body={email.body} anchors={anchors} serveDate={serveDate} onChange={(body) => set({ body })} />
      </section>

      <section className="space-y-3">
      <StepTitle n={2} title="Connect earlier facts" note="Only add a dependency if this email relies on a previous email." />
      <DependencyPicker email={email} node={node} onChange={(depends_on) => set({ depends_on })} />
      </section>

      <section className="space-y-3">
      <StepTitle n={3} title="Tell the grader the perfect answer" note="This is the answer key, not text the model sees." />
      <AnswerKeyBuilder answer={email.answer} anchors={anchors} serveDate={serveDate} obligations={obligationNames(node)} obligationKinds={obligationKinds(node)} reuseAnchors={anchors} bodyAnchors={ownBodyAnchors} anchorOrigins={origins} onChange={(answer) => set({ answer })} />
      </section>
    </div>
  );
}

// The only sender control: one dropdown for "who is this email from?". Every email in the benchmark
// is written TO the CEO (the boss the assistant works for), so there is no recipient picker — To is
// fixed and shown read-only. Picking a sender records `from`, pins `to: CEO`, and quietly materializes
// the person into this storyline's cast (cast is inbox dressing, never graded) so the export stays
// readable. Choices: you (the CEO) for self-notes, anyone already in this storyline, then the rest of
// the standard roster, plus "Someone else…" for a one-off name.
function SenderPicker({ node, email, onUpdateNode }: { node: CorpusNode; email: Email; onUpdateNode: (n: CorpusNode) => void }) {
  const cast = node.cast ?? {};
  const seen = new Set<string>();
  const options: { key: string; label: string }[] = [];
  const add = (key: string, label: string) => { if (!seen.has(key)) { seen.add(key); options.push({ key, label }); } };
  add("CEO", "you (the CEO)");
  if (email.from) add(email.from, cast[email.from] || email.from);   // keep the current sender selectable
  for (const [k, name] of Object.entries(cast)) add(k, name || k);
  for (const p of STANDARD_ROSTER) add(p.key, p.name);

  function pick(key: string, name: string) {
    const nextCast: Record<string, string> = { ...cast, CEO: cast.CEO ?? "you" };
    if (nextCast[key] === undefined) nextCast[key] = name;
    onUpdateNode({ ...node, cast: nextCast, emails: node.emails.map((e) => e.id === email.id ? { ...e, from: key, to: "CEO" } : e) });
  }
  function onChange(v: string) {
    if (v === "__custom__") {
      const raw = window.prompt('Who is this email from? Type a name or role, e.g. "Acme account manager".');
      if (!raw) return;
      const key = normalizeKey(raw); if (!key) return;
      pick(key, normalizeName(raw));
      return;
    }
    const p = options.find((o) => o.key === v); if (p) pick(p.key, cast[p.key] ?? p.label);
  }

  return (
    <div className="space-y-2">
      <label className="block text-xs text-slate-400">From
        <select value={email.from || ""} onChange={(e) => onChange(e.target.value)}
          className="mt-1 w-full rounded border border-slate-700 bg-slate-800 px-2 py-1.5 text-sm text-slate-100">
          {!email.from && <option value="">pick who it&apos;s from…</option>}
          {options.map((o) => <option key={o.key} value={o.key}>{o.label}</option>)}
          <option value="__custom__">Someone else…</option>
        </select>
      </label>
      <div className="flex items-center gap-2 text-xs text-slate-400">To
        <span className="rounded border border-slate-700 bg-slate-800 px-2 py-1 text-slate-300">you (the CEO)</span>
        <span className="text-slate-500">every email is written to the boss the assistant works for</span>
      </div>
    </div>
  );
}

function StepTitle({ n, title, note }: { n: number; title: string; note: string }) {
  return (
    <div className="mb-3 flex items-start gap-2">
      <span className="grid h-6 w-6 shrink-0 place-items-center rounded-full bg-sky-600 text-xs font-semibold text-white">{n}</span>
      <div>
        <h2 className="text-sm font-semibold text-slate-100">{title}</h2>
        <p className="text-xs text-slate-500">{note}</p>
      </div>
    </div>
  );
}

function DependencyPicker({ email, node, onChange }: { email: Email; node: CorpusNode; onChange: (edges: Edge[]) => void }) {
  // Only offer OTHER emails in THIS storyline. A prerequisite from an unrelated node would just be
  // confusing noise to an author scoped to one storyline (cross-node edges are an advanced, rarely-used
  // case; any that already exist in the data still render below as their raw id / @node).
  const others = node.emails.filter((e) => e.id !== email.id);
  const chosen = new Set(email.depends_on.map((d) => d.email).filter(Boolean) as string[]);
  // The deadline distinction (static vs date edge) is an advanced detail most authors never touch:
  // new edges start as `static` (plain "comes after"), and needles auto-upgrade to `date` via
  // withDerivedDateEdges. So we hide the type dropdown behind an Advanced toggle and just show the
  // current setting as a muted word; opening Advanced reveals the editable selects.
  const [showTypes, setShowTypes] = useState(false);
  function add(emailId: string) {
    if (!emailId || chosen.has(emailId)) return;
    onChange([...email.depends_on, { email: emailId, type: "static" }]);
  }
  function setType(emailId: string, type: EdgeType) {
    onChange(email.depends_on.map((d) => d.email === emailId ? { ...d, type } : d));
  }

  return (
    <div className="rounded-lg border border-slate-800 bg-slate-900/40 p-3">
      <h3 className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-400">Depends on</h3>
      <div className="space-y-1.5">
        {email.depends_on.map((d, i) => (
          <div key={i} className="flex items-center gap-2 text-xs">
            <code className="flex-1 truncate rounded bg-slate-800 px-2 py-1 font-mono text-slate-300">{d.email ?? `@node:${d.node}`}</code>
            {showTypes ? (
              <select value={d.type} onChange={(e) => d.email && setType(d.email, e.target.value as EdgeType)}
                className="rounded border border-slate-700 bg-slate-800 px-2 py-1 text-slate-200">
                <option value="static">comes after (no deadline)</option>
                <option value="date">comes after, with a deadline</option>
              </select>
            ) : (
              <span className="text-slate-500" title="Most links just mean this email comes later. Open Advanced to set a deadline.">{d.type === "date" ? "deadline" : "comes after"}</span>
            )}
            <button onClick={() => onChange(email.depends_on.filter((_, j) => j !== i))} className="px-1 text-slate-500 hover:text-rose-400">✕</button>
          </div>
        ))}
        <select value="" onChange={(e) => add(e.target.value)} className="w-full rounded border border-slate-700 bg-slate-800 px-2 py-1 text-xs text-slate-200">
          <option value="">+ add a prerequisite email…</option>
          {others.filter((e) => !chosen.has(e.id)).map((e) => <option key={e.id} value={e.id}>{e.id} · {e.subject}</option>)}
        </select>
        {email.depends_on.length > 0 && (
          <button type="button" onClick={() => setShowTypes((s) => !s)} className="text-[11px] text-slate-500 hover:text-sky-300">
            {showTypes ? "▾ hide deadline settings" : "▸ Advanced: set deadlines"}
          </button>
        )}
      </div>
    </div>
  );
}
