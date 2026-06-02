import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import { api } from "./api";
import { DEFAULT_SCENARIO } from "./types";
import type {
  AnswerForm,
  Chip,
  EmailForm,
  Graph,
  OracleResult,
  ThreadForm,
} from "./types";

// Distinct scenario names present in the graph, plus any just-created empty ones,
// with the default bucket sorted last. Drives the tab bar.
function scenarioList(graph: Graph | null, pending: string[]): string[] {
  const names = new Set<string>(pending);
  graph?.threads.forEach((t) => names.add(t.scenario || DEFAULT_SCENARIO));
  if (names.size === 0) names.add(DEFAULT_SCENARIO);
  return [...names].sort((a, b) =>
    a === DEFAULT_SCENARIO ? 1 : b === DEFAULT_SCENARIO ? -1 : a.localeCompare(b)
  );
}

// All threads transitively RELATED to `threadId` — via depends_on, node_depends_on,
// or a shared anchor/fact reference. This is exactly the set that must travel
// together to keep scenarios isolated, so moving a thread moves its whole group.
function relatedThreads(graph: Graph, threadId: string): Set<string> {
  const threadOf = (eid: string) => graph.emails[eid]?.thread;
  const adj: Record<string, Set<string>> = {};
  const link = (a?: string, b?: string) => {
    if (!a || !b || a === b) return;
    (adj[a] ||= new Set()).add(b);
    (adj[b] ||= new Set()).add(a);
  };
  Object.values(graph.emails).forEach((e) => {
    e.depends_on.forEach((d) => d.email && link(e.thread, threadOf(d.email)));
    const ans = e.answer;
    ans.expect.forEach((ex) => {
      [ex.duration, typeof ex.count === "string" ? ex.count : null].forEach((v) => {
        if (typeof v === "string") {
          for (const m of v.matchAll(/@([A-Za-z_]\w*)/g)) {
            const src = graph.fact_map?.[m[1]] || graph.emission_map?.[m[1]];
            if (src) link(e.thread, threadOf(src));
          }
        }
      });
    });
  });
  graph.threads.forEach((t) =>
    (t.node_depends_on || []).forEach((d) => d.node && link(t.id, d.node))
  );

  const seen = new Set<string>([threadId]);
  const stack = [threadId];
  while (stack.length) {
    const cur = stack.pop()!;
    (adj[cur] || new Set()).forEach((n) => {
      if (!seen.has(n)) { seen.add(n); stack.push(n); }
    });
  }
  return seen;
}

const LAYOUT_KEY = "sb.layout";

function loadLayout(): Record<string, { x: number; y: number }> {
  try {
    return JSON.parse(localStorage.getItem(LAYOUT_KEY) || "{}");
  } catch {
    return {};
  }
}
function saveLayout(l: Record<string, { x: number; y: number }>) {
  localStorage.setItem(LAYOUT_KEY, JSON.stringify(l));
}

// deterministic fallback layout: lay each thread out in a column
function autoLayout(graph: Graph, saved: Record<string, { x: number; y: number }>) {
  graph.threads.forEach((t, ti) => {
    t.emails.forEach((eid, ei) => {
      const e = graph.emails[eid];
      if (!e) return;
      const pos = saved[eid];
      e.x = pos ? pos.x : 80 + ti * 340;
      e.y = pos ? pos.y : 80 + ei * 160;
    });
  });
}

interface Store {
  graph: Graph | null;
  selected: string | null;
  dirty: Set<string>;
  oracle: OracleResult | null;
  busy: boolean;
  scenarios: string[];
  activeScenario: string;
  setActiveScenario: (name: string) => void;
  createScenario: () => string;
  renameScenario: (oldName: string, name: string) => void;
  moveThreadToScenario: (threadId: string, name: string) => void;
  reload: () => Promise<void>;
  select: (id: string | null) => void;
  setPosition: (id: string, x: number, y: number) => void;
  updateEmail: (id: string, patch: Partial<EmailForm>) => void;
  updateAnswer: (id: string, answer: AnswerForm) => void;
  updateThread: (id: string, patch: Partial<ThreadForm>) => void;
  newEmail: () => void;
  addFollowUp: (sourceId: string) => void;
  deleteEmail: (id: string) => void;
  deleteThread: (id: string) => Promise<void>;
  connect: (sourceId: string, targetId: string) => void;
  disconnect: (targetId: string, sourceId: string) => void;
  ensureDateEdge: (emailId: string, anchorName: string) => void;
  save: () => Promise<void>;
  runOracle: () => Promise<void>;
}

const Ctx = createContext<Store | null>(null);
export const useStore = () => {
  const s = useContext(Ctx);
  if (!s) throw new Error("useStore outside provider");
  return s;
};

export function StoreProvider({ children }: { children: React.ReactNode }) {
  const [graph, setGraph] = useState<Graph | null>(null);
  const [selected, setSelected] = useState<string | null>(null);
  const [dirty, setDirty] = useState<Set<string>>(new Set());
  const [oracle, setOracle] = useState<OracleResult | null>(null);
  const [busy, setBusy] = useState(false);
  const [activeScenario, setActiveScenario] = useState<string>(DEFAULT_SCENARIO);
  const [pendingScenarios, setPendingScenarios] = useState<string[]>([]);
  const layout = useRef(loadLayout());

  const reload = useCallback(async () => {
    setBusy(true);
    try {
      const g = await api.corpus();
      autoLayout(g, layout.current);
      setGraph(g);
      setDirty(new Set());
      // keep the current tab if it still has threads; otherwise fall back to the first.
      const names = scenarioList(g, []);
      setActiveScenario((cur) => (names.includes(cur) ? cur : names[0]));
      setPendingScenarios((p) => p.filter((name) => !names.includes(name)));
    } finally {
      setBusy(false);
    }
  }, []);

  useEffect(() => {
    reload();
  }, [reload]);

  const markDirty = useCallback((threadId: string) => {
    setDirty((d) => new Set(d).add(threadId));
  }, []);

  const mutate = useCallback((fn: (g: Graph) => void) => {
    setGraph((prev) => {
      if (!prev) return prev;
      const next: Graph = { ...prev, emails: { ...prev.emails }, threads: [...prev.threads] };
      fn(next);
      return next;
    });
  }, []);

  const select = useCallback((id: string | null) => setSelected(id), []);

  const setPosition = useCallback((id: string, x: number, y: number) => {
    layout.current[id] = { x, y };
    saveLayout(layout.current);
    mutate((g) => {
      const e = g.emails[id];
      if (e) g.emails[id] = { ...e, x, y };
    });
  }, [mutate]);

  const updateEmail = useCallback((id: string, patch: Partial<EmailForm>) => {
    mutate((g) => {
      const e = g.emails[id];
      if (!e) return;
      g.emails[id] = { ...e, ...patch };
      markDirty(e.thread);
    });
  }, [mutate, markDirty]);

  const updateAnswer = useCallback((id: string, answer: AnswerForm) => {
    updateEmail(id, { answer });
  }, [updateEmail]);

  const updateThread = useCallback((id: string, patch: Partial<ThreadForm>) => {
    mutate((g) => {
      g.threads = g.threads.map((t) => (t.id === id ? { ...t, ...patch } : t));
      markDirty(id);
    });
  }, [mutate, markDirty]);

  // --- scenarios (the editor's tabs) ---------------------------------------

  // Create an empty scenario and switch to it. It only exists client-side until
  // the first email is added to it (then it's a real thread that gets saved).
  const createScenario = useCallback((): string => {
    const existing = new Set(scenarioList(graph, pendingScenarios));
    let n = existing.size;
    let name = `scenario-${n}`;
    while (existing.has(name)) name = `scenario-${++n}`;
    setPendingScenarios((p) => [...p, name]);
    setActiveScenario(name);
    return name;
  }, [graph, pendingScenarios]);

  // Reassign every thread in a scenario to a new name — used to rename a tab and,
  // crucially, to MERGE two tabs (rename A's threads into B). Pure relabel; no
  // relations change, so a merge of two valid scenarios stays valid.
  const renameScenario = useCallback((oldName: string, name: string) => {
    if (!name || oldName === name) return;
    mutate((g) => {
      g.threads = g.threads.map((t) => {
        if (t.scenario !== oldName) return t;
        markDirty(t.id);
        return { ...t, scenario: name };
      });
    });
    setPendingScenarios((p) => p.map((s) => (s === oldName ? name : s)));
    setActiveScenario((cur) => (cur === oldName ? name : cur));
  }, [mutate, markDirty]);

  // Move a thread AND everything related to it (its relation-component) into the
  // target scenario — so a thread never gets separated from the emails it depends
  // on or shares a value with, which would otherwise break scenario isolation.
  const moveThreadToScenario = useCallback((threadId: string, name: string) => {
    if (!name) return;
    mutate((g) => {
      const group = relatedThreads(g, threadId);
      g.threads = g.threads.map((t) => (group.has(t.id) ? { ...t, scenario: name } : t));
      group.forEach(markDirty);
    });
    setActiveScenario(name);
  }, [mutate, markDirty]);

  // a brand-new conversation: one thread, one email, placed to the right of the field
  const newEmail = useCallback(() => {
    if (!graph) return;
    let n = 1;
    let id = `conversation-${n}`;
    while (graph.threads.some((t) => t.id === id)) id = `conversation-${++n}`;
    const eid = `${id}.email1`;
    const maxX = Math.max(40, ...Object.values(graph.emails).map((e) => e.x ?? 0));
    const x = Object.keys(graph.emails).length ? maxX + 340 : 120;
    mutate((g) => {
      g.threads = [...g.threads, { id, cast: { ME: "you", THEM: "" }, scenario: activeScenario, node_depends_on: [], emails: [eid] }];
      g.emails[eid] = blankEmail(eid, id, x, 120);
      markDirty(id);
    });
    setSelected(eid);
  }, [graph, mutate, markDirty, activeScenario]);

  // a follow-up to an existing email: same thread, linked, placed below it
  const addFollowUp = useCallback((sourceId: string) => {
    if (!graph) return;
    const src = graph.emails[sourceId];
    if (!src) return;
    const threadId = src.thread;
    const t = graph.threads.find((x) => x.id === threadId);
    if (!t) return;
    let n = t.emails.length + 1;
    let eid = `${threadId}.email${n}`;
    while (graph.emails[eid]) eid = `${threadId}.email${++n}`;
    const x = src.x ?? 80;
    const y = (src.y ?? 40) + 170;
    mutate((g) => {
      g.threads = g.threads.map((x) => (x.id === threadId ? { ...x, emails: [...x.emails, eid] } : x));
      const e = blankEmail(eid, threadId, x, y);
      e.from = src.from;
      e.to = src.to;
      e.depends_on = [{ email: sourceId, type: "static" }];
      g.emails[eid] = e;
      markDirty(threadId);
    });
    setSelected(eid);
  }, [graph, mutate, markDirty]);

  const deleteEmail = useCallback((id: string) => {
    mutate((g) => {
      const e = g.emails[id];
      if (!e) return;
      const tid = e.thread;
      delete g.emails[id];
      // drop edges pointing at it
      Object.keys(g.emails).forEach((k) => {
        const em = g.emails[k];
        if (em.depends_on.some((d) => d.email === id)) {
          g.emails[k] = { ...em, depends_on: em.depends_on.filter((d) => d.email !== id) };
          markDirty(em.thread);
        }
      });
      g.threads = g.threads.map((t) => (t.id === tid ? { ...t, emails: t.emails.filter((x) => x !== id) } : t));
      markDirty(tid);
    });
    setSelected((s) => (s === id ? null : s));
  }, [mutate, markDirty]);

  const deleteThread = useCallback(async (id: string) => {
    setBusy(true);
    try {
      await api.deleteThread(id);
      await reload();
      setSelected(null);
    } finally {
      setBusy(false);
    }
  }, [reload]);

  const connect = useCallback((sourceId: string, targetId: string) => {
    mutate((g) => {
      const t = g.emails[targetId];
      const s = g.emails[sourceId];
      if (!t || !s || sourceId === targetId) return;
      // A relation can't cross scenarios. If it would, MERGE: pull every thread of
      // the source's scenario into the target's (a pure relabel — both were valid,
      // so the union is too). This is the "merge tabs on link" behaviour.
      const sThread = g.threads.find((x) => x.id === s.thread);
      const tThread = g.threads.find((x) => x.id === t.thread);
      if (sThread && tThread && sThread.scenario !== tThread.scenario) {
        const from = sThread.scenario;
        const into = tThread.scenario;
        const moved = g.threads.filter((x) => x.scenario === from).map((x) => x.id);
        g.threads = g.threads.map((x) => (x.scenario === from ? { ...x, scenario: into } : x));
        moved.forEach(markDirty);
      }
      if (t.depends_on.some((d) => d.email === sourceId)) return;
      g.emails[targetId] = { ...t, depends_on: [...t.depends_on, { email: sourceId, type: "static" }] };
      markDirty(t.thread);
    });
  }, [mutate, markDirty]);

  const disconnect = useCallback((targetId: string, sourceId: string) => {
    mutate((g) => {
      const t = g.emails[targetId];
      if (!t) return;
      g.emails[targetId] = { ...t, depends_on: t.depends_on.filter((d) => d.email !== sourceId) };
      markDirty(t.thread);
    });
  }, [mutate, markDirty]);

  const ensureDateEdge = useCallback((emailId: string, anchorName: string) => {
    mutate((g) => {
      const src = g.emission_map[anchorName];
      if (!src || src === emailId) return;
      const e = g.emails[emailId];
      if (!e) return;
      const existing = e.depends_on.find((d) => d.email === src);
      if (existing) {
        if (existing.type !== "date")
          g.emails[emailId] = {
            ...e,
            depends_on: e.depends_on.map((d) => (d.email === src ? { ...d, type: "date" } : d)),
          };
      } else {
        g.emails[emailId] = { ...e, depends_on: [...e.depends_on, { email: src, type: "date" }] };
      }
      markDirty(e.thread);
    });
  }, [mutate, markDirty]);

  const save = useCallback(async () => {
    if (!graph) return;
    setBusy(true);
    try {
      for (const tid of dirty) {
        const t = graph.threads.find((x) => x.id === tid);
        if (!t) continue;
        const payload = {
          id: t.id,
          cast: t.cast,
          scenario: t.scenario,
          node_depends_on: t.node_depends_on,
          emails: t.emails.map((eid) => graph.emails[eid]).filter(Boolean),
        };
        await api.putThread(payload);
      }
      await reload();
    } finally {
      setBusy(false);
    }
  }, [graph, dirty, reload]);

  const runOracle = useCallback(async () => {
    setBusy(true);
    try {
      setOracle(await api.oracle());
    } finally {
      setBusy(false);
    }
  }, []);

  const scenarios = useMemo(() => scenarioList(graph, pendingScenarios), [graph, pendingScenarios]);

  const value = useMemo<Store>(() => ({
    graph, selected, dirty, oracle, busy,
    scenarios, activeScenario, setActiveScenario, createScenario, renameScenario, moveThreadToScenario,
    reload, select, setPosition, updateEmail, updateAnswer, updateThread,
    newEmail, addFollowUp, deleteEmail, deleteThread, connect, disconnect,
    ensureDateEdge, save, runOracle,
  }), [graph, selected, dirty, oracle, busy, scenarios, activeScenario,
       createScenario, renameScenario, moveThreadToScenario,
       reload, select, setPosition,
       updateEmail, updateAnswer, updateThread, newEmail, addFollowUp, deleteEmail,
       deleteThread, connect, disconnect, ensureDateEdge, save, runOracle]);

  return <Ctx.Provider value={value}>{children}</Ctx.Provider>;
}

function blankEmail(id: string, thread: string, x: number, y: number): EmailForm {
  return {
    id, thread, from: "", to: ["ME"], subject: "",
    body_segments: [{ type: "text", value: "" }],
    depends_on: [], answer: { expect: [], forbid: [], emits: {}, facts: {} },
    emits: {}, reachable_anchors: [], defined_facts: {}, reachable_facts: [], x, y,
  };
}
