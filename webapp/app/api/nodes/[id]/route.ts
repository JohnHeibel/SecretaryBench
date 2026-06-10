import { NextResponse } from "next/server";
import { getNode, upsertNode, deleteNode } from "@/lib/store";
import type { CorpusNode } from "@/lib/types";

export const dynamic = "force-dynamic";

type Ctx = { params: Promise<{ id: string }> };

export async function GET(_req: Request, { params }: Ctx) {
  const { id } = await params;
  const node = await getNode(id);
  return node ? NextResponse.json(node) : NextResponse.json({ error: "not found" }, { status: 404 });
}

export async function PUT(req: Request, { params }: Ctx) {
  const { id } = await params;
  const node = (await req.json()) as CorpusNode;
  if (node.id !== id) return NextResponse.json({ error: "id mismatch" }, { status: 400 });
  const by = req.headers.get("x-author") ?? "anon";
  const saved = await upsertNode(node, by);
  if (!saved) return NextResponse.json({ error: "this storyline belongs to another author" }, { status: 403 });
  return NextResponse.json(saved);
}

export async function DELETE(req: Request, { params }: Ctx) {
  const { id } = await params;
  const by = req.headers.get("x-author") ?? "anon";
  if (!(await deleteNode(id, by))) return NextResponse.json({ error: "only the author who created this storyline can delete it" }, { status: 403 });
  return NextResponse.json({ deleted: id });
}
