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
import type {
  AnswerForm,
  Chip,
  EmailForm,
  Graph,
  OracleResult,
  ThreadForm,
} from "./types";

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
  reload: () => Promise<void>;
  select: (id: string | null) => void;
  setPosition: (id: string, x: number, y: number) => void;
  updateEmail: (id: string, patch: Partial<EmailForm>) => void;
  updateAnswer: (id: string, answer: AnswerForm) => void;
  updateThread: (id: string, patch: Partial<ThreadForm>) => void;
  addThread: () => void;
  addEmail: (threadId: string) => void;
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
  const layout = useRef(loadLayout());

  const reload = useCallback(async () => {
    setBusy(true);
    try {
      const g = await api.corpus();
      autoLayout(g, layout.current);
      setGraph(g);
      setDirty(new Set());
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

  const addThread = useCallback(() => {
    mutate((g) => {
      let n = g.threads.length + 1;
      let id = `thread-${n}`;
      while (g.threads.some((t) => t.id === id)) id = `thread-${++n}`;
      const eid = `${id}.email1`;
      g.threads = [...g.threads, { id, cast: { ME: "you" }, node_depends_on: [], emails: [eid] }];
      g.emails[eid] = blankEmail(eid, id, 80 + g.threads.length * 340, 80);
      markDirty(id);
    });
  }, [mutate, markDirty]);

  const addEmail = useCallback((threadId: string) => {
    if (!graph) return;
    const t = graph.threads.find((x) => x.id === threadId);
    if (!t) return;
    let n = t.emails.length + 1;
    let eid = `${threadId}.email${n}`;
    while (graph.emails[eid]) eid = `${threadId}.email${++n}`;
    const last = t.emails.length ? graph.emails[t.emails[t.emails.length - 1]] : null;
    const x = last?.x ?? 80;
    const y = (last?.y ?? 40) + 160;
    mutate((g) => {
      g.threads = g.threads.map((x) => (x.id === threadId ? { ...x, emails: [...x.emails, eid] } : x));
      g.emails[eid] = blankEmail(eid, threadId, x, y);
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
      if (!t || sourceId === targetId) return;
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

  const value = useMemo<Store>(() => ({
    graph, selected, dirty, oracle, busy,
    reload, select, setPosition, updateEmail, updateAnswer, updateThread,
    addThread, addEmail, deleteEmail, deleteThread, connect, disconnect,
    ensureDateEdge, save, runOracle,
  }), [graph, selected, dirty, oracle, busy, reload, select, setPosition,
       updateEmail, updateAnswer, updateThread, addThread, addEmail, deleteEmail,
       deleteThread, connect, disconnect, ensureDateEdge, save, runOracle]);

  return <Ctx.Provider value={value}>{children}</Ctx.Provider>;
}

function blankEmail(id: string, thread: string, x: number, y: number): EmailForm {
  return {
    id, thread, from: "", to: ["ME"], subject: "",
    body_segments: [{ type: "text", value: "" }],
    depends_on: [], answer: { expect: [], forbid: [], emits: {} },
    emits: {}, reachable_anchors: [], x, y,
  };
}
