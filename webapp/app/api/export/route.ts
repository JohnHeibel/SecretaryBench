import JSZip from "jszip";
import { listNodes } from "@/lib/store";

export const dynamic = "force-dynamic";

// GET /api/export -> a zip of nodes/<id>.json, byte-for-byte what
// sb.schema.load_corpus expects. Drop the contents into corpus/ to run the
// benchmark on the app-authored corpus.
export async function GET() {
  const nodes = await listNodes();
  const zip = new JSZip();
  const dir = zip.folder("nodes")!;
  for (const node of nodes) dir.file(`${node.id}.json`, JSON.stringify(node, null, 2) + "\n");
  const blob = await zip.generateAsync({ type: "arraybuffer" });
  return new Response(blob, {
    headers: {
      "content-type": "application/zip",
      "content-disposition": `attachment; filename="corpus-${new Date().toISOString().slice(0, 10)}.zip"`,
    },
  });
}
