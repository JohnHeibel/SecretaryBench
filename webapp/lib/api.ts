// Client-side fetchers. CRUD hits same-origin Next routes. The validation
// endpoints (/api/resolve, /api/lint) are the Python functions — same origin on
// Vercel, or a standalone dev validator via NEXT_PUBLIC_VALIDATOR_BASE locally.
import type { CorpusNode, LintResult, OracleResult, ResolveResult } from "./types";

const VALIDATOR = process.env.NEXT_PUBLIC_VALIDATOR_BASE ?? "";

// Anonymous per-browser identity — no login. A random token minted once per browser and sent as
// x-author on every store call; the server stamps it as owner on first save and requires it on delete.
// The coordinator can become admin by setting this key to the deployment's SB_ADMIN_TOKEN value:
//   localStorage.setItem("sb-author", "<token>")  (then reload)
export function authorToken(): string {
  if (typeof window === "undefined") return "anon";
  let t = window.localStorage.getItem("sb-author");
  if (!t) { t = crypto.randomUUID(); window.localStorage.setItem("sb-author", t); }
  return t;
}

export async function fetchNodes(): Promise<{ nodes: CorpusNode[]; mine: Set<string> }> {
  const r = await fetch("/api/nodes", { cache: "no-store", headers: { "x-author": authorToken() } });
  const { nodes, mine } = (await r.json()) as { nodes: CorpusNode[]; mine: string[] };
  return { nodes, mine: new Set(mine) };
}

export async function saveNode(node: CorpusNode): Promise<CorpusNode> {
  const r = await fetch(`/api/nodes/${encodeURIComponent(node.id)}`, {
    method: "PUT",
    headers: { "content-type": "application/json", "x-author": authorToken() },
    body: JSON.stringify(node),
  });
  if (r.status === 403) throw new Error("not-owner");
  if (!r.ok) throw new Error(`save failed: ${r.status}`);
  return r.json();
}

// Resolves true iff the server actually deleted the row (false = not this browser's storyline).
export async function deleteNode(id: string): Promise<boolean> {
  const r = await fetch(`/api/nodes/${encodeURIComponent(id)}`, { method: "DELETE", headers: { "x-author": authorToken() } });
  return r.ok;
}

export async function resolveExpr(expr: string, serveDate: string, anchors: Record<string, string> = {}): Promise<ResolveResult> {
  try {
    const r = await fetch(`${VALIDATOR}/api/resolve`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ expr, serve_date: serveDate, anchors }),
    });
    return r.json();
  } catch {
    return { ok: false, error: "validator unreachable (is the Python function running?)" };
  }
}

export async function lintCorpus(nodes: CorpusNode[]): Promise<LintResult> {
  try {
    const r = await fetch(`${VALIDATOR}/api/lint`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ nodes }),
    });
    return r.json();
  } catch {
    return { ok: false, error: "validator unreachable (is the Python function running?)" };
  }
}

export async function oracleCorpus(nodes: CorpusNode[]): Promise<OracleResult> {
  try {
    const r = await fetch(`${VALIDATOR}/api/oracle`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ nodes }),
    });
    return r.json();
  } catch {
    return { ok: false, error: "validator unreachable (is the Python function running?)" };
  }
}
