// Server-side corpus store. Two backends behind one interface:
//   - Postgres (Neon on Vercel) when POSTGRES_URL is set — one `nodes` table,
//     each row a JSONB document mirroring a corpus/nodes/*.json file. Seeded once
//     from the bundled snapshot when the table is empty (a fresh Neon DB).
//   - A local JSON file (.data/nodes.json) otherwise, so `npm run dev` works with
//     zero setup.
// Both seed from `seed/nodes.json`, a build-time snapshot of the git corpus
// produced by scripts/vendor_sb.py. It is imported (not fs-read from ../corpus)
// so it ships inside the Vercel bundle — the git corpus is not on that filesystem.
import fs from "node:fs";
import path from "node:path";
import seedNodes from "@/seed/nodes.json";
import type { CorpusNode } from "./types";

// Neon's native Vercel integration injects DATABASE_URL; the older Vercel Postgres
// product (and @vercel/postgres, which reads it) uses POSTGRES_URL. Alias so either
// integration variant works — otherwise prod silently falls back to the file backend
// on a read-only serverless FS and every save fails.
if (!process.env.POSTGRES_URL && process.env.DATABASE_URL) {
  process.env.POSTGRES_URL = process.env.DATABASE_URL;
}

const usePg = !!process.env.POSTGRES_URL;
const SEED = seedNodes as unknown as CorpusNode[];

// --- file backend (local dev) ---------------------------------------------

const DATA_DIR = path.join(process.cwd(), ".data");
const DATA_FILE = path.join(DATA_DIR, "nodes.json");

function fileReadAll(): Record<string, CorpusNode> {
  if (!fs.existsSync(DATA_FILE)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
    const seed: Record<string, CorpusNode> = {};
    for (const node of SEED) seed[node.id] = node;
    fs.writeFileSync(DATA_FILE, JSON.stringify(seed, null, 2));
    return seed;
  }
  return JSON.parse(fs.readFileSync(DATA_FILE, "utf8"));
}

function fileWriteAll(all: Record<string, CorpusNode>): void {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  fs.writeFileSync(DATA_FILE, JSON.stringify(all, null, 2));
}

// --- postgres backend (prod) ----------------------------------------------

async function pg() {
  const { sql } = await import("@vercel/postgres");
  return sql;
}

let pgReady: Promise<void> | null = null;
async function ensurePg(): Promise<void> {
  if (!pgReady) {
    pgReady = (async () => {
      const sql = await pg();
      await sql`CREATE TABLE IF NOT EXISTS nodes (
        id text PRIMARY KEY,
        data jsonb NOT NULL,
        updated_at timestamptz NOT NULL DEFAULT now(),
        updated_by text
      )`;
      // Seed a fresh (empty) Neon DB from the bundled corpus snapshot, so the app
      // opens with the corpus instead of a blank slate. ON CONFLICT DO NOTHING means
      // it is a one-time fill that never clobbers edited rows on later boots (and is
      // safe if two cold-start invocations race). EXISTS returns a boolean, dodging
      // the int8-as-string quirk that would make a count(*) === 0 check silently fail.
      const { rows } = await sql`SELECT EXISTS (SELECT 1 FROM nodes) AS has_rows`;
      if (!rows[0]?.has_rows) {
        for (const node of SEED) {
          await sql`INSERT INTO nodes (id, data, updated_by) VALUES (${node.id}, ${JSON.stringify(node)}::jsonb, ${"seed"})
            ON CONFLICT (id) DO NOTHING`;
        }
      }
    })();
  }
  return pgReady;
}

// --- public interface ------------------------------------------------------

export async function listNodes(): Promise<CorpusNode[]> {
  if (usePg) {
    await ensurePg();
    const sql = await pg();
    const { rows } = await sql`SELECT data FROM nodes ORDER BY id`;
    return rows.map((r) => r.data as CorpusNode);
  }
  return Object.values(fileReadAll()).sort((a, b) => a.id.localeCompare(b.id));
}

export async function getNode(id: string): Promise<CorpusNode | null> {
  if (usePg) {
    await ensurePg();
    const sql = await pg();
    const { rows } = await sql`SELECT data FROM nodes WHERE id = ${id}`;
    return rows[0] ? (rows[0].data as CorpusNode) : null;
  }
  return fileReadAll()[id] ?? null;
}

export async function upsertNode(node: CorpusNode, by = "anon"): Promise<CorpusNode> {
  if (usePg) {
    await ensurePg();
    const sql = await pg();
    const json = JSON.stringify(node);
    await sql`INSERT INTO nodes (id, data, updated_by) VALUES (${node.id}, ${json}::jsonb, ${by})
      ON CONFLICT (id) DO UPDATE SET data = EXCLUDED.data, updated_by = EXCLUDED.updated_by, updated_at = now()`;
    return node;
  }
  const all = fileReadAll();
  all[node.id] = node;
  fileWriteAll(all);
  return node;
}

// Bulk ingest, the mirror of /api/export. `mode: "upsert"` (default) merges the
// incoming nodes over what's already stored; `mode: "replace"` wipes the store first
// so the corpus becomes exactly the imported set. Postgres does it in one round-trip
// (a single multi-row INSERT for the batch, plus a TRUNCATE up front when replacing)
// rather than N awaits; the file backend reads once, mutates the map, writes once.
export async function importNodes(nodes: CorpusNode[], by = "anon", mode: "upsert" | "replace" = "upsert"): Promise<number> {
  if (usePg) {
    await ensurePg();
    const sql = await pg();
    if (mode === "replace") await sql`TRUNCATE nodes`;
    if (nodes.length) {
      const ids = nodes.map((n) => n.id);
      const datas = nodes.map((n) => JSON.stringify(n));
      const bys = nodes.map(() => by);
      // unnest the three parallel arrays into rows, then one upsert for the whole batch
      await sql`INSERT INTO nodes (id, data, updated_by)
        SELECT id, data::jsonb, updated_by
        FROM unnest(${ids as any}::text[], ${datas as any}::text[], ${bys as any}::text[]) AS t(id, data, updated_by)
        ON CONFLICT (id) DO UPDATE SET data = EXCLUDED.data, updated_by = EXCLUDED.updated_by, updated_at = now()`;
    }
    return nodes.length;
  }
  const all = mode === "replace" ? {} : fileReadAll();
  for (const node of nodes) all[node.id] = node;
  fileWriteAll(all);
  return nodes.length;
}

export async function deleteNode(id: string): Promise<void> {
  if (usePg) {
    await ensurePg();
    const sql = await pg();
    await sql`DELETE FROM nodes WHERE id = ${id}`;
    return;
  }
  const all = fileReadAll();
  delete all[id];
  fileWriteAll(all);
}
