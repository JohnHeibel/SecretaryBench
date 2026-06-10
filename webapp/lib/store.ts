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
        updated_by text,
        owner text
      )`;
      // Ownership column for tables created before the 2026-06-10 incident. `owner` is the anonymous
      // per-browser token that first saved the row; only that browser (or the admin token) may delete it.
      await sql`ALTER TABLE nodes ADD COLUMN IF NOT EXISTS owner text`;
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

// --- ownership --------------------------------------------------------------
// No login, on purpose. The client mints a random token per browser (localStorage) and sends it as
// `x-author`. The first save of a storyline stamps it as the row's owner; only a matching token may
// delete the row. SB_ADMIN_TOKEN (env) overrides everything. Tokens written by header-less tools
// ("anon") or the seeder ("seed") are stored as NULL = unowned = admin-only delete.
const UNOWNED = new Set(["", "anon", "seed"]);
const realOwner = (by: string): string | null => (UNOWNED.has(by) ? null : by);
export const isAdmin = (token: string | null): boolean => !!token && !!process.env.SB_ADMIN_TOKEN && token === process.env.SB_ADMIN_TOKEN;

// Which node ids the viewer may delete (drives the UI's delete buttons; the DELETE route re-checks).
export async function ownedIds(viewer: string): Promise<string[]> {
  if (usePg) {
    await ensurePg();
    const sql = await pg();
    const { rows } = isAdmin(viewer) ? await sql`SELECT id FROM nodes` : await sql`SELECT id FROM nodes WHERE owner = ${viewer}`;
    return rows.map((r) => r.id as string);
  }
  return Object.keys(fileReadAll()); // local dev is single-user: everything is yours
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
    // A brand-new row is stamped with its creator's token; an update never reassigns ownership.
    await sql`INSERT INTO nodes (id, data, updated_by, owner) VALUES (${node.id}, ${json}::jsonb, ${by}, ${realOwner(by)})
      ON CONFLICT (id) DO UPDATE SET data = EXCLUDED.data, updated_by = EXCLUDED.updated_by, updated_at = now()`;
    return node;
  }
  const all = fileReadAll();
  all[node.id] = node;
  fileWriteAll(all);
  return node;
}

// Returns true iff the row was deleted. Only the owner token (or admin) may delete; unowned rows
// (owner NULL: pre-incident rows, seed, header-less writes) are admin-only.
export async function deleteNode(id: string, by = "anon"): Promise<boolean> {
  if (usePg) {
    await ensurePg();
    const sql = await pg();
    const { rowCount } = isAdmin(by)
      ? await sql`DELETE FROM nodes WHERE id = ${id}`
      : realOwner(by) ? await sql`DELETE FROM nodes WHERE id = ${id} AND owner = ${by}` : { rowCount: 0 };
    return (rowCount ?? 0) > 0;
  }
  const all = fileReadAll();
  delete all[id];
  fileWriteAll(all);
  return true;
}
