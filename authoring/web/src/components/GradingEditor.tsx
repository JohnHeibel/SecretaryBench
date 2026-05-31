import { emptyChip } from "../chipLabel";
import type { AnswerForm, Chip, EmailForm, ExpectForm, Predicate } from "../types";
import { ChipPill } from "./ChipPill";

interface Props {
  email: EmailForm;
  serve: string;
  anchors: Record<string, string>;
  onChange: (answer: AnswerForm) => void;
  onAnchorPicked: (name: string) => void;
}

const ACTIONS: { value: ExpectForm["action"]; label: string }[] = [
  { value: "create_event", label: "Calendar event" },
  { value: "create_todo", label: "To-do" },
  { value: "reschedule", label: "Reschedule event" },
];

const MATCHES: { value: Predicate["match"]; label: string }[] = [
  { value: "on", label: "exactly on" },
  { value: "by", label: "by (deadline)" },
  { value: "within", label: "within window" },
  { value: "any_of", label: "any of" },
];

const TOLERANCES = [
  { value: "exact_day", label: "same day" },
  { value: "exact_time", label: "exact day & time" },
  { value: "within:1d", label: "± 1 day" },
  { value: "within:3d", label: "± 3 days" },
];

export function GradingEditor({ email, serve, anchors, onChange, onAnchorPicked }: Props) {
  const answer = email.answer;
  const setExpect = (i: number, patch: Partial<ExpectForm>) =>
    onChange({ ...answer, expect: answer.expect.map((e, j) => (j === i ? { ...e, ...patch } : e)) });

  const addExpect = () =>
    onChange({
      ...answer,
      expect: [...answer.expect, { action: "create_event", title_match: [], when: { match: "on", chip: emptyChip("next_weekday") }, tolerance: "exact_day", count: 1 }],
    });

  const removeExpect = (i: number) =>
    onChange({ ...answer, expect: answer.expect.filter((_, j) => j !== i) });

  const noAction = answer.expect.length === 0 && answer.forbid.length === 0;

  const chipChanged = (c: Chip) => {
    if (c.base.kind === "anchor" && c.base.name) onAnchorPicked(c.base.name);
  };

  return (
    <div>
      <div className="mb-2 flex items-center justify-between">
        <h3 className="text-sm font-semibold text-white/80">Correct answer</h3>
        <button onClick={addExpect} className="rounded-md border border-edge bg-panel2 px-2 py-1 text-xs text-accent hover:border-accent">
          + expected action
        </button>
      </div>

      {noAction && (
        <div className="rounded-lg border border-edge/60 bg-panel/50 p-3 text-sm text-muted">
          No expected actions — this email is graded as <b className="text-white/80">“correctly take no action”</b>. Add an
          expected action above if the secretary should schedule something.
        </div>
      )}

      <div className="space-y-3">
        {answer.expect.map((e, i) => {
          const isTodo = e.action === "create_todo";
          return (
            <div key={i} className="rounded-lg border border-edge bg-panel2/60 p-3">
              <div className="mb-2 flex items-center gap-2">
                <select className={inp} value={e.action} onChange={(ev) => setExpect(i, { action: ev.target.value as any })}>
                  {ACTIONS.map((a) => <option key={a.value} value={a.value}>{a.label}</option>)}
                </select>
                <button onClick={() => removeExpect(i)} className="ml-auto text-xs text-muted hover:text-bad">remove</button>
              </div>

              <Field label="title contains">
                <input
                  className={inp}
                  placeholder="keywords, comma-separated (e.g. acme, renewal)"
                  value={e.title_match.join(", ")}
                  onChange={(ev) => setExpect(i, { title_match: ev.target.value.split(",").map((s) => s.trim()).filter(Boolean) })}
                />
              </Field>

              <Field label={isTodo ? "due" : "when"}>
                <PredicatePicker
                  predicate={e.when}
                  onChange={(when) => setExpect(i, { when })}
                  serve={serve}
                  anchors={anchors}
                  reachableAnchors={email.reachable_anchors}
                  onChipChange={chipChanged}
                />
              </Field>

              <div className="flex flex-wrap gap-3">
                {!isTodo && (
                  <Field label="length" inline>
                    <input className={`${inp} w-24`} placeholder="e.g. 90m" value={e.duration || ""}
                      onChange={(ev) => setExpect(i, { duration: ev.target.value || null })} />
                  </Field>
                )}
                <Field label="how many" inline>
                  <input type="number" className={`${inp} w-20`} value={e.count ?? ""} placeholder="any"
                    onChange={(ev) => setExpect(i, { count: ev.target.value === "" ? null : Number(ev.target.value) })} />
                </Field>
                <Field label="tolerance" inline>
                  <select className={inp} value={e.tolerance} onChange={(ev) => setExpect(i, { tolerance: ev.target.value })}>
                    {TOLERANCES.map((t) => <option key={t.value} value={t.value}>{t.label}</option>)}
                  </select>
                </Field>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function PredicatePicker({
  predicate, onChange, serve, anchors, reachableAnchors, onChipChange,
}: {
  predicate: Predicate | null;
  onChange: (p: Predicate) => void;
  serve: string;
  anchors: Record<string, string>;
  reachableAnchors: string[];
  onChipChange: (c: Chip) => void;
}) {
  const p = predicate || { match: "on", chip: emptyChip("next_weekday") };
  const setChip = (c: Chip) => { onChipChange(c); onChange({ ...p, chip: c }); };

  return (
    <div className="flex flex-wrap items-center gap-2">
      <select
        className={inp}
        value={p.match}
        onChange={(e) => {
          const match = e.target.value as Predicate["match"];
          if (match === "any_of") onChange({ match, chips: p.chips || [p.chip || emptyChip("next_weekday")] });
          else onChange({ match, chip: p.chip || emptyChip("next_weekday"), ...(match === "within" ? { avoid_chip: p.avoid_chip } : {}) });
        }}
      >
        {MATCHES.map((m) => <option key={m.value} value={m.value}>{m.label}</option>)}
      </select>

      {p.match === "any_of" ? (
        <div className="flex flex-wrap items-center gap-1">
          {(p.chips || []).map((c, i) => (
            <ChipPill
              key={i}
              chip={c}
              onChange={(nc) => { onChipChange(nc); onChange({ ...p, chips: (p.chips || []).map((x, j) => (j === i ? nc : x)) }); }}
              onRemove={() => onChange({ ...p, chips: (p.chips || []).filter((_, j) => j !== i) })}
              serve={serve} anchors={anchors} reachableAnchors={reachableAnchors}
            />
          ))}
          <button className="rounded border border-edge px-1.5 text-xs text-accent"
            onClick={() => onChange({ ...p, chips: [...(p.chips || []), emptyChip("next_weekday")] })}>+</button>
        </div>
      ) : (
        <ChipPill chip={p.chip || emptyChip("next_weekday")} onChange={setChip}
          serve={serve} anchors={anchors} reachableAnchors={reachableAnchors} />
      )}

      {p.match === "within" && (
        <span className="flex items-center gap-1 text-xs text-muted">
          avoiding
          {p.avoid_chip ? (
            <ChipPill chip={p.avoid_chip} onChange={(c) => onChange({ ...p, avoid_chip: c })}
              onRemove={() => onChange({ ...p, avoid_chip: undefined })}
              serve={serve} anchors={anchors} reachableAnchors={reachableAnchors} />
          ) : (
            <button className="rounded border border-edge px-1.5 text-accent"
              onClick={() => onChange({ ...p, avoid_chip: emptyChip("next_weekday") })}>+ blackout</button>
          )}
        </span>
      )}
    </div>
  );
}

function Field({ label, children, inline }: { label: string; children: React.ReactNode; inline?: boolean }) {
  return (
    <div className={inline ? "mb-2" : "mb-2"}>
      <div className="mb-1 text-[11px] uppercase tracking-wide text-muted">{label}</div>
      {children}
    </div>
  );
}

const inp =
  "rounded-md border border-edge bg-ink px-2 py-1 text-sm text-white outline-none focus:border-accent";
