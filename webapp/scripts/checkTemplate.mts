// Guard for the loadable example: serialize the template to a corpus and run it through the REAL
// reference solver (sb.oracle via python3), asserting it lints clean AND solves 1.0. "Project Atlas" is
// the one thing a new author loads, so a satisfiability regression in it would teach the wrong thing.
// Exits non-zero if any email doesn't pass. Run: npx tsx scripts/checkTemplate.mts
import { execFileSync } from "node:child_process";
import { projectAtlasNode } from "../lib/templates.ts";
import type { CorpusNode } from "../lib/types.ts";

const code = [
  'import sys, json',
  'sys.path.insert(0, "..")',                          // webapp/ is cwd; ".." is the repo root with sb/
  'from datetime import date',
  'from sb import schema, engine',
  'from sb.oracle import oracle_model',
  'from sb.scheduler import build_plan',
  'nodes = json.load(sys.stdin)',
  'corpus = schema.build_corpus(nodes)',               // raises CorpusError if not well-formed
  'days = max(200, len(corpus.emails) + 120)',         // Atlas reaches ~12 weeks out; give the plan room
  'plan = build_plan(corpus, start_date=date(2026,6,1), seed=42, n_days=days)',
  'res = engine.run(corpus, plan, oracle_model)',
  'print("SCORE", res.score(), "passed", res.passed, "total", res.total)',
  'for eid, r in res.results.items(): print(("PASS " if r.passed else "FAIL ") + eid + ("" if r.passed else " -> " + str(r.headline)))',
  'sys.exit(0 if res.score() == 1.0 else 1)',
].join("\n");

const templates: Array<{ name: string; node: CorpusNode }> = [
  { name: "project_atlas", node: projectAtlasNode() },
];

let failed = false;
for (const { name, node } of templates) {
  try {
    process.stdout.write(execFileSync("python3", ["-c", code], { cwd: process.cwd(), input: JSON.stringify([node]), encoding: "utf8" }));
    console.log(`${name} template solves 1.0 ✓`);
  } catch (e: any) {
    process.stdout.write(String(e?.stdout ?? ""));
    console.error(`✗ ${name} template did NOT solve 1.0 (or errored):\n` + String(e?.stderr ?? e));
    failed = true;
  }
}
if (failed) process.exit(1);
