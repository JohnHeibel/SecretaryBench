// Client-side fetchers. CRUD hits same-origin Next routes. The validation
// endpoints (/api/resolve, /api/lint) are the Python functions — same origin on
// Vercel, or a standalone dev validator via NEXT_PUBLIC_VALIDATOR_BASE locally.
import type { CorpusNode, LintResult, OracleResult, ResolveResult } from "./types";

const VALIDATOR = process.env.NEXT_PUBLIC_VALIDATOR_BASE ?? "";

export async function fetchNodes(): Promise<CorpusNode[]> {
  const r = await fetch("/api/nodes", { cache: "no-store" });
  return r.json();
}

export async function saveNode(node: CorpusNode): Promise<CorpusNode> {
  const r = await fetch(`/api/nodes/${encodeURIComponent(node.id)}`, {
    method: "PUT",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(node),
  });
  if (!r.ok) throw new Error(`save failed: ${r.status}`);
  return r.json();
}

export async function deleteNode(id: string): Promise<void> {
  await fetch(`/api/nodes/${encodeURIComponent(id)}`, { method: "DELETE" });
}

// Bulk import — the mirror of the /api/export download. Posts a corpus and returns
// how many nodes landed. Accepts either nodes parsed from JSON or a raw export .zip
// (Blob); the route sniffs the content-type. `mode: "replace"` makes the corpus
// exactly the imported set, default "upsert" overlays it.
export async function importCorpus(payload: CorpusNode[] | Blob, mode: "upsert" | "replace" = "upsert"): Promise<{ imported: number; mode: string }> {
  const zip = payload instanceof Blob;
  const r = await fetch(`/api/import?mode=${mode}`, {
    method: "POST",
    headers: { "content-type": zip ? "application/zip" : "application/json" },
    body: zip ? payload : JSON.stringify({ nodes: payload }),
  });
  if (!r.ok) throw new Error((await r.json().catch(() => ({}))).error ?? `import failed: ${r.status}`);
  return r.json();
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
