import { NextResponse } from "next/server";
import { listNodes, ownedIds, upsertNode } from "@/lib/store";
import type { CorpusNode } from "@/lib/types";

export const dynamic = "force-dynamic";

// With an x-author header (the webapp client): { nodes, mine } where `mine` lists the ids this
// browser may delete. Without one (sb.sync, curl, the rescue script): the plain node array, unchanged.
export async function GET(req: Request) {
  const viewer = req.headers.get("x-author");
  const nodes = await listNodes();
  if (!viewer) return NextResponse.json(nodes);
  return NextResponse.json({ nodes, mine: await ownedIds(viewer) });
}

export async function POST(req: Request) {
  const node = (await req.json()) as CorpusNode;
  if (!node?.id) return NextResponse.json({ error: "node needs an id" }, { status: 400 });
  const by = req.headers.get("x-author") ?? "anon";
  const saved = await upsertNode(node, by);
  if (!saved) return NextResponse.json({ error: "this storyline belongs to another author" }, { status: 403 });
  return NextResponse.json(saved);
}
