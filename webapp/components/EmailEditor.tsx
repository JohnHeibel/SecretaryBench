"use client";
import { useEffect, useState } from "react";
import type { CorpusNode, Edge, EdgeType, Email } from "@/lib/types";
import { anchorOrigins, bodyAnchorNames, obligationKinds, obligationNames } from "@/lib/grammar";
import { MAX_PERSON_NAME, STANDARD_ROSTER, normalizeName } from "@/lib/people";
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

function castKeys(node: CorpusNode): string[] {
  return Object.keys(node.cast ?? {});
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
      <CastManager node={node} onUpdateNode={onUpdateNode} />
      {email ? (
        <EmailPanel node={node} email={email} allNodes={allNodes} anchors={anchors} serveDate={serveDate} onUpdateEmail={onUpdateEmail} onAutoSlugEmail={onAutoSlugEmail} />
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
      <p className="mt-1 text-slate-400">You&apos;re inside a <strong className="text-slate-300">storyline</strong>, a group of related emails. The <strong className="text-slate-300">cast</strong> is the people in it. It starts with <span className="font-mono text-sky-300">CEO → you</span>, the person the AI works for.</p>
      <ol className="ml-4 mt-1 list-decimal space-y-0.5">
        <li>Write the email (who it&apos;s from, the subject, the body). Use <span className="font-mono text-sky-300">+ insert date</span> for any date so it stays exact.</li>
        <li>In <strong>Answer key</strong>, say what to do: name the thing, pick its date, or tick <em>&ldquo;this email needs no action.&rdquo;</em></li>
        <li>The bar at the bottom should read <span className="text-emerald-400">Ready for export</span> and <span className="text-emerald-400">oracle solves 100%</span>. That means a perfect assistant could actually do it.</li>
      </ol>
      <p className="mt-1 text-slate-400">Want more? <a href="/guide" className="text-sky-400 underline hover:text-sky-300">Open the full walkthrough →</a> (worked examples, the date builder, how to build a needle)</p>
    </div>
  );
}

function CastManager({ node, onUpdateNode }: { node: CorpusNode; onUpdateNode: (n: CorpusNode) => void }) {
  const cast = node.cast ?? {};
  const entries = Object.entries(cast);
  const present = new Set(Object.keys(cast));
  const [open, setOpen] = useState(entries.length < 2);

  function setCast(next: Record<string, string>) { onUpdateNode({ ...node, cast: next }); }

  function addPerson(key: string, name: string) {
    if (cast[key] !== undefined) return;        // already in this node's cast
    setCast({ ...cast, [key]: name });
  }
  function removePerson(key: string) {
    const used = node.emails.some((e) => e.from === key || asList(e.to).includes(key) || asList(e.cc).includes(key));
    const msg = used
      ? `"${key}" appears in one or more emails (From / To / Cc). Remove them from the cast anyway?`
      : `Remove "${key}" from the cast?`;
    if (!window.confirm(msg)) return;
    const c = { ...cast }; delete c[key]; setCast(c);
  }

  const available = STANDARD_ROSTER.filter((p) => !present.has(p.key));

  return (
    <div className="rounded-lg border border-slate-800 bg-slate-900/40">
      <button onClick={() => setOpen(!open)} className="flex w-full items-center justify-between px-3 py-2 text-xs font-semibold uppercase tracking-wide text-slate-400">
        <span>Cast for {node.id} <span className="font-normal lowercase text-slate-500">({entries.length} people)</span></span>
        <span>{open ? "▾" : "▸"}</span>
      </button>
      {open && (
        <div className="space-y-1.5 border-t border-slate-800 p-3">
          <p className="text-xs leading-5 text-slate-500">Add everyone who can appear in From, To, or Cc. Pick from the <strong>standard roster</strong> so the same role is spelled the same way across storylines. Keep <code className="font-mono text-sky-300">CEO</code> as the assistant&apos;s boss.</p>
          {entries.map(([key, name]) => (
            <div key={key} className="flex items-center gap-2">
              <code className="w-32 rounded border border-slate-700 bg-slate-800 px-2 py-1 text-xs text-slate-400 select-all">{key}</code>
              <input value={name} maxLength={MAX_PERSON_NAME}
                onChange={(e) => setCast({ ...cast, [key]: e.target.value.slice(0, MAX_PERSON_NAME) })}
                onBlur={(e) => setCast({ ...cast, [key]: normalizeName(e.target.value) })}
                placeholder="Display name (Role)" className="flex-1 rounded border border-slate-700 bg-slate-800 px-2 py-1 text-xs text-slate-200" />
              <button onClick={() => removePerson(key)} className="px-1 text-xs text-slate-500 hover:text-rose-400">✕</button>
            </div>
          ))}
          <div className="flex flex-wrap items-center gap-2 pt-0.5">
            <select value="" onChange={(e) => { const p = STANDARD_ROSTER.find((x) => x.key === e.target.value); if (p) addPerson(p.key, p.name); }}
              disabled={available.length === 0}
              className="rounded border border-slate-700 bg-slate-800 px-2 py-1 text-xs text-slate-200 disabled:opacity-40">
              <option value="">{available.length ? "+ add from standard roster…" : "all standard people added"}</option>
              {available.map((p) => <option key={p.key} value={p.key}>{p.key} · {p.name}</option>)}
            </select>
          </div>
        </div>
      )}
    </div>
  );
}

// A stored recipient field is a string OR a list (or missing). Normalize to a list so the
// multi-recipient picker has one shape to work with; a legacy single string shows as one chip.
function asList(v: string | string[] | undefined): string[] {
  if (Array.isArray(v)) return v.filter(Boolean);
  return v ? [v] : [];
}

function EmailPanel({ node, email, allNodes, anchors, serveDate, onUpdateEmail, onAutoSlugEmail }: Omit<Props, "onUpdateNode" | "onRenameNode" | "onAddEmail"> & { email: Email }) {
  const keys = castKeys(node);
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

      <div className="space-y-2">
        <label className="block text-xs text-slate-400 sm:w-1/2 sm:pr-1.5">From
          <select value={email.from} onChange={(e) => set({ from: e.target.value })}
            className="mt-1 w-full rounded border border-slate-700 bg-slate-800 px-2 py-1.5 text-sm text-slate-100">
            <option value="">pick a person</option>
            {keys.map((k) => <option key={k} value={k}>{k} · {node.cast[k]}</option>)}
          </select>
        </label>
        <RecipientPicker label="To" hint="who the email is sent to" cast={node.cast} keys={keys}
          selected={asList(email.to)} onChange={(to) => set({ to })} />
        <RecipientPicker label="Cc" hint="copied on the email (optional); the assistant can see this" cast={node.cast} keys={keys}
          selected={asList(email.cc)} onChange={(cc) => set({ cc })} />
      </div>

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

// A dense multi-recipient picker over the node's cast: tap a chip to add/remove that person.
// Stores the selection as an array (so To/Cc can hold several people). Used for both To and Cc;
// From stays a single select. Empty selection is fine for Cc.
function RecipientPicker({ label, hint, cast, keys, selected, onChange }: {
  label: string; hint: string; cast: Record<string, string>; keys: string[]; selected: string[]; onChange: (v: string[]) => void;
}) {
  const chosen = new Set(selected);
  const toggle = (k: string) => onChange(chosen.has(k) ? selected.filter((s) => s !== k) : [...selected, k]);
  // A selected person no longer in the cast (renamed/removed) still shows so it's never silently lost.
  const extra = selected.filter((s) => !keys.includes(s));
  return (
    <div className="text-xs text-slate-400">
      <span title={hint}>{label}</span>
      {keys.length === 0 && extra.length === 0 ? (
        <p className="mt-1 text-slate-500">Add people to the cast above, then pick recipients here.</p>
      ) : (
        <div className="mt-1 flex flex-wrap gap-1.5">
          {keys.map((k) => {
            const on = chosen.has(k);
            return (
              <button key={k} type="button" onClick={() => toggle(k)} title={cast[k] || k}
                className={`rounded-full border px-2.5 py-1 text-xs transition-colors ${on ? "border-sky-500 bg-sky-600/20 text-sky-200" : "border-slate-700 bg-slate-800 text-slate-300 hover:border-slate-500"}`}>
                {on ? "✓ " : ""}{k}
              </button>
            );
          })}
          {extra.map((k) => (
            <button key={k} type="button" onClick={() => toggle(k)} title="not in the cast; click to remove"
              className="rounded-full border border-amber-600/70 bg-amber-900/20 px-2.5 py-1 text-xs text-amber-300">✓ {k} ⚠</button>
          ))}
        </div>
      )}
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
