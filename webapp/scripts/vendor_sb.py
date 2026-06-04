#!/usr/bin/env python3
"""Vendor the keystone sb modules into the webapp so the Vercel Python functions
can import them. This is a PURE COPY — no rewriting — so the resolver/schema the
web app validates against is byte-identical to what the benchmark grades with.
Drift between the two is the exact bug this whole app exists to prevent, so we
copy rather than reimplement.

Copies repo `sb/{__init__,resolver,schema,grader,engine,scheduler,oracle}.py` ->
`webapp/api/_lib/sb/`, and snapshots `corpus/nodes/*.json` -> `webapp/seed/nodes.json`
so a fresh Postgres/Neon deploy can seed itself (the git corpus is not on the
Vercel filesystem; only `webapp/` ships). Run automatically by `npm run dev` /
`npm run build`; also runnable by hand.
"""
from __future__ import annotations

import filecmp
import json
import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve()
WEBAPP = HERE.parent.parent           # webapp/
REPO = WEBAPP.parent                   # repo root
SRC = REPO / "sb"
DST = WEBAPP / "api" / "_lib" / "sb"
CORPUS = REPO / "corpus" / "nodes"
SEED = WEBAPP / "seed" / "nodes.json"

# The pure-stdlib subset the validation functions need. resolver.py is
# stdlib-only; schema/grader/engine/scheduler/oracle import only each other and
# resolver; __init__ is a docstring. /api/oracle needs the full solve pipeline
# (oracle -> engine+scheduler -> grader -> schema -> resolver).
MODULES = ["__init__.py", "resolver.py", "schema.py", "grader.py",
           "engine.py", "scheduler.py", "oracle.py"]


def snapshot_corpus() -> int:
    """Bundle the git corpus into the webapp as a single JSON array, so store.ts
    can seed an empty Postgres/Neon DB on first boot (and local dev) from a file
    that actually ships to Vercel. Stable key order keeps the snapshot diff-clean."""
    SEED.parent.mkdir(parents=True, exist_ok=True)
    nodes = [json.loads(f.read_text()) for f in sorted(CORPUS.glob("*.json"))] if CORPUS.exists() else []
    SEED.write_text(json.dumps(sorted(nodes, key=lambda n: n["id"]), indent=2, sort_keys=True) + "\n")
    return len(nodes)


GEN_TS = WEBAPP / "lib" / "schema.generated.ts"


def generate_ts_types() -> None:
    """Emit lib/schema.generated.ts from sb.schema's grammar constants, so the answer-key
    types the editor uses cannot drift from the grammar the grader enforces (the ADR-0001
    bug, where the UI kept the dead expect/forbid model while the backend moved to ops/verbs).
    Reads the VENDORED copy (always present after the copy/fallback in main), and is
    deterministic so the committed file stays diff-clean. A build never fails on codegen — if
    sb can't be imported we keep the committed file and warn."""
    try:
        sys.path.insert(0, str(DST.parent))            # api/_lib, so `import sb.schema` resolves
        import importlib
        schema = importlib.import_module("sb.schema")
        importlib.reload(schema)
        union = lambda vals: " | ".join(f'"{v}"' for v in vals)
        verb_keys = "\n".join(f"  {v}?: string;" for v in schema.VERB_ORDER)
        pred_keys = "\n".join(f"  {k}?: {'string[]' if k == 'any_of' else 'string'};" for k in schema.PREDICATE_OPS)
        ts = f'''// AUTO-GENERATED from sb/schema.py by scripts/vendor_sb.py — DO NOT EDIT BY HAND.
// Regenerated on every `npm run vendor` / `npm run build`. To change the answer-key grammar,
// edit the Python grammar lists (VERB_ORDER / KIND_ORDER / EDGE_ORDER / PREDICATE_OPS) and
// rebuild; the editor then fails typecheck if it still uses a field the grammar dropped.

export type Verb = {union(schema.VERB_ORDER)};
export type ObjKind = {union(schema.KIND_ORDER)};
export type EdgeType = {union(schema.EDGE_ORDER)};

// Date predicate over an answer-key slot. Exactly one key is set; any_of takes a list.
export interface Predicate {{
{pred_keys}
}}

// One verb on a named obligation. Serializes 1:1 to corpus JSON: the verb is the KEY and its
// value is the obligation name, e.g. {{ "create": "kickoff", "kind": "event", "on": {{...}} }}.
export interface Op {{
{verb_keys}
  kind?: ObjKind;
  on?: Predicate;
  match?: string[];
  tolerance?: string;
}}

export interface Answer {{
  ops: Op[];
  emits?: Record<string, string>;
}}
'''
        GEN_TS.write_text(ts)
        print(f"generated {GEN_TS.relative_to(WEBAPP)} from sb.schema")
    except Exception as exc:                            # never block a build on codegen
        print(f"WARN: could not regenerate {GEN_TS.name} ({exc}); keeping committed copy")


def check() -> int:
    """Anti-drift gate: fail (nonzero) if the vendored copy differs from sb/. The
    whole app exists to validate with the SAME code the grader uses, so a stale
    vendored copy is a silent correctness bug. Run in CI: `vendor_sb.py --check`."""
    drift = [name for name in MODULES
             if not (DST / name).exists() or not filecmp.cmp(SRC / name, DST / name, shallow=False)]
    if drift:
        print(f"DRIFT: vendored sb differs from sb/ for: {', '.join(drift)} — run `npm run vendor`")
        return 1
    print(f"vendored sb is in sync with sb/ ({len(MODULES)} modules)")
    return 0


def main() -> None:
    if "--check" in sys.argv[1:]:
        raise SystemExit(check())
    if not SRC.exists():
        # A CLI/archive deploy uploads only webapp/, so the parent sb/ and corpus/
        # aren't here. Fall back to the copy that was vendored before upload — but
        # only if it is COMPLETE, else fail loudly rather than ship stale/missing code.
        missing = [n for n in MODULES if not (DST / n).exists()]
        if missing or not SEED.exists():
            raise SystemExit(f"cannot find sb/ at {SRC} and the vendored copy is incomplete "
                             f"(missing: {missing or ['seed/nodes.json']}) — run `npm run vendor` from a full checkout")
        print(f"sb/ not present (archive deploy) — using the pre-vendored {len(MODULES)} modules + seed already in webapp/")
        generate_ts_types()
        return
    DST.mkdir(parents=True, exist_ok=True)
    for name in MODULES:
        src = SRC / name
        if not src.exists():
            raise SystemExit(f"missing expected module {src}")
        shutil.copy2(src, DST / name)
    n = snapshot_corpus()
    print(f"vendored {len(MODULES)} sb modules -> {DST.relative_to(REPO)}; seeded {n} nodes -> {SEED.relative_to(REPO)}")
    generate_ts_types()


if __name__ == "__main__":
    main()
