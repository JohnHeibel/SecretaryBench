import { NextResponse } from "next/server";
import { listNodes, upsertNode } from "@/lib/store";
import type { CorpusNode } from "@/lib/types";

export const dynamic = "force-dynamic";

export async function GET() {
  return NextResponse.json(await listNodes());
}

export async function POST(req: Request) {
  const node = (await req.json()) as CorpusNode;
  if (!node?.id) return NextResponse.json({ error: "node needs an id" }, { status: 400 });
  const by = req.headers.get("x-author") ?? "anon";
  return NextResponse.json(await upsertNode(node, by));
}
