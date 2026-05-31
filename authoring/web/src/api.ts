import type {
  Chip,
  Graph,
  OracleResult,
  ResolveResult,
  Segment,
  ThreadForm,
} from "./types";

const TOKEN = localStorage.getItem("authToken") || "";

async function req<T>(path: string, opts: RequestInit = {}): Promise<T> {
  const res = await fetch(`/api${path}`, {
    ...opts,
    headers: {
      "Content-Type": "application/json",
      ...(TOKEN ? { "X-Auth-Token": TOKEN } : {}),
      ...(opts.headers || {}),
    },
  });
  if (!res.ok) {
    const detail = await res.text();
    throw new Error(`${res.status}: ${detail}`);
  }
  return res.json();
}

export const api = {
  corpus: () => req<Graph>("/corpus"),

  putThread: (thread: { id: string; cast: Record<string, string>; node_depends_on: any[]; emails: any[] }) =>
    req<{ errors: string[] }>(`/thread/${thread.id}`, {
      method: "PUT",
      body: JSON.stringify(thread),
    }),

  deleteThread: (id: string) =>
    req<{ errors: string[] }>(`/thread/${id}`, { method: "DELETE" }),

  validate: () => req<{ errors: string[]; sample: Graph["sample"] }>("/validate", { method: "POST" }),

  resolve: (chip: Chip | null, serve: string, anchors: Record<string, string>, expr?: string) =>
    req<ResolveResult>("/resolve", {
      method: "POST",
      body: JSON.stringify({ chip, expr, serve, anchors }),
    }),

  render: (segments: Segment[], serve: string, anchors: Record<string, string>) =>
    req<{ ok: boolean; text?: string; emissions?: Record<string, string>; error?: string }>("/render", {
      method: "POST",
      body: JSON.stringify({ body_segments: segments, serve, anchors }),
    }),

  oracle: () => req<OracleResult>("/oracle", { method: "POST" }),
};
