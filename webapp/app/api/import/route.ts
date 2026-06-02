import { NextResponse } from "next/server";
import JSZip from "jszip";
import { importNodes } from "@/lib/store";
import type { CorpusNode } from "@/lib/types";

export const dynamic = "force-dynamic";

// POST /api/import — the mirror of GET /api/export. Accepts either
//   - an export .zip (content-type application/zip or a multipart upload), whose
//     nodes/<id>.json entries are read back exactly as export wrote them, or
//   - a JSON body { nodes: CorpusNode[] }.
// Bulk-upserts the nodes into the store in one shot. Merge semantics: ?mode=upsert
// (default) overlays the imported nodes on the existing corpus; ?mode=replace wipes
// the store first so the corpus becomes exactly the imported set.
async function nodesFromZip(buf: ArrayBuffer): Promise<CorpusNode[]> {
  const zip = await JSZip.loadAsync(buf);
  const out: CorpusNode[] = [];
  const files = Object.values(zip.files).filter((f) => !f.dir && /(^|\/)nodes\/.+\.json$/.test(f.name));
  for (const f of files) out.push(JSON.parse(await f.async("string")) as CorpusNode);
  return out;
}

export async function POST(req: Request) {
  const mode = new URL(req.url).searchParams.get("mode") === "replace" ? "replace" : "upsert";
  const by = req.headers.get("x-author") ?? "anon";
  const ctype = req.headers.get("content-type") ?? "";

  let nodes: CorpusNode[];
  try {
    if (ctype.includes("application/zip") || ctype.includes("application/octet-stream")) {
      nodes = await nodesFromZip(await req.arrayBuffer());
    } else if (ctype.includes("multipart/form-data")) {
      const file = (await req.formData()).get("file");
      if (!(file instanceof Blob)) return NextResponse.json({ error: "no 'file' in upload" }, { status: 400 });
      nodes = await nodesFromZip(await file.arrayBuffer());
    } else {
      const body = (await req.json()) as { nodes?: CorpusNode[] };
      if (!Array.isArray(body?.nodes)) return NextResponse.json({ error: "expected { nodes: [] }" }, { status: 400 });
      nodes = body.nodes;
    }
  } catch (e) {
    return NextResponse.json({ error: `could not parse corpus: ${(e as Error).message}` }, { status: 400 });
  }

  const bad = nodes.find((n) => !n?.id || !Array.isArray(n.emails));
  if (bad) return NextResponse.json({ error: "every node needs an id and an emails[]" }, { status: 400 });

  const imported = await importNodes(nodes, by, mode);
  return NextResponse.json({ imported, mode });
}
