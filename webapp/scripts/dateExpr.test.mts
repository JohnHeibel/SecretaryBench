// Throwaway verification for lib/dateExpr.ts: (1) every canonical grammar string parses,
// (2) round-trips serialize(parse(s)) === s, and (3) the serialized string is accepted by the
// REAL Python resolver (sb/resolver.py via `python3 -c`). Run: npx tsx scripts/dateExpr.test.mts
import { execFileSync } from "node:child_process";
import { parseExpr, serializeExpr } from "../lib/dateExpr.ts";

// Canonical strings spanning every production (mirrors ANSWER_KEY_GRAMMAR.md §3 examples).
const CANON = [
  "serve", "serve+9d", "serve+3bd", "serve-2w", "serve+1m", "serve+1y",
  "@signing", "@signing+2w", "@blackout",
  "next:THU", "this:FRI", "next:MON from @signing", "this:FRI from @migration+1w",
  "next:WED from serve+2d",
  "nth:3,FRI,+1m", "nth:last,FRI,0m", "nth:1,MON,-1m",
  "dom:25,0m", "dom:1,+2m",
  "week_of:(serve+1w)", "week_of:(@signing)", "month:0m", "month:+1m", "month:-2m",
  "week_of:(next:MON from @signing)",
];

// Strings the resolver accepts but that aren't our canonical serialization — must still parse,
// and re-serialize to a semantically-equal canonical form (checked via the resolver below).
// `+9d` canonicalizes to `serve+9d`; `month:1m` to `month:+1m` (explicit base / signed monthref).
const NONCANON = ["serve +5d", "+9d", "+3bd", "month:1m"];

// Strings the parser must REJECT (return null) so the UI falls back to raw mode. `1m` is a bare
// monthref — the real resolver rejects it as a top-level expr too, so we must as well.
const REJECT = ["", "garbage", "serve+5", "next:THURSDAY", "week_of:(serve", "@", "nth:9", "1m"];

let fail = 0;
const note = (ok: boolean, msg: string) => { if (!ok) { fail++; console.error("  ✗ " + msg); } };

// resolve a string with the real Python resolver; returns {ok, human|error}
function pyResolve(expr: string): { ok: boolean; out: string } {
  const code = `import sys; sys.path.insert(0, "..");\n` +
    `from sb import resolver\n` +
    `from datetime import date\n` +
    `ctx = resolver.Context(serve=date(2026,6,1), anchors={"signing": date(2026,6,8), "blackout": resolver.Interval(date(2026,8,3), date(2026,8,9)), "migration": date(2026,7,15)})\n` +
    `try:\n  v = resolver.resolve(${JSON.stringify(expr)}, ctx); print("OK", resolver.human(v))\n` +
    `except Exception as e:\n  print("ERR", e)`;
  try {
    const out = execFileSync("python3", ["-c", code], { cwd: process.cwd(), encoding: "utf8" }).trim();
    return { ok: out.startsWith("OK"), out };
  } catch (e: any) { return { ok: false, out: String(e?.stderr ?? e) }; }
}

console.log("1) canonical: parse + exact round-trip + resolver accepts");
for (const s of CANON) {
  const e = parseExpr(s);
  note(e !== null, `parse failed: ${s}`);
  if (e) note(serializeExpr(e) === s, `round-trip drift: ${s} -> ${serializeExpr(e)}`);
  const r = pyResolve(s);
  note(r.ok, `resolver rejected canonical: ${s} (${r.out})`);
}

console.log("2) non-canonical: parses + re-serializes to a resolver-equal string");
for (const s of NONCANON) {
  const e = parseExpr(s);
  note(e !== null, `parse failed (noncanon): ${s}`);
  if (e) {
    const reser = serializeExpr(e);
    const a = pyResolve(s), b = pyResolve(reser);
    note(a.ok && b.ok && a.out === b.out, `semantic drift: ${s} (${a.out}) vs ${reser} (${b.out})`);
  }
}

console.log("3) reject: parser returns null (-> raw escape hatch)");
for (const s of REJECT) note(parseExpr(s) === null, `should reject but parsed: ${JSON.stringify(s)} -> ${JSON.stringify(parseExpr(s))}`);

console.log(fail === 0 ? "\nALL PASS ✓" : `\n${fail} FAILURES ✗`);
process.exit(fail === 0 ? 0 : 1);
