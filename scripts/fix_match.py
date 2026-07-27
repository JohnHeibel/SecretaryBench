"""
Give each obligation a model-robust `match` keyword set derived from the op NAME
(generic scheduling words dropped; least-shared content token(s) kept so it stays
unique within the scenario). Only `match` changes — never op name/kind/date.

Genuinely-ambiguous sibling pairs (one obligation's identity ⊆ another's) can't be
auto-distinguished; the linter flags them and we revert those to their ORIGINAL match
and report them for a human authoring fix. Gate: linter + oracle == 100%.
"""
import json, re, glob, sys
from pathlib import Path
from datetime import date
from collections import Counter
from sb.schema import load_corpus, build_corpus, CorpusError
from sb.scheduler import build_plan, Levers
from sb.engine import Store, run
from sb.oracle import oracle_model

# Reads the recovered corpus in corpus/nodes and rewrites it in place. Run AFTER
# scripts/recover_corpus.py. Originals are read into memory first, so in-place is safe.
NODE_DIR = Path("corpus/nodes")
BK = NODE_DIR
STOP = set("""the a an of to for and or with on in out up at by is are be been being this that
these those your you our their his her its it whether can could should would will shall may
might must do does did done have has had not no yes please need needs want make made create
creates created contact inform pick fill decide schedule set send get put let know talk about
discuss discussion meeting meet sync call review session event todo task reminder note update
follow followup plan planning arrange organize confirm check ensure prepare added new final go
day date time week weeks month""".split())

def toks(name):
    parts = re.findall(r"[A-Za-z0-9]+", name)
    out = []
    for p in parts:
        out += re.findall(r"[A-Z]+(?=[A-Z][a-z])|[A-Z]?[a-z']+|[A-Z]+|[0-9]+", p) or [p]
    return [t.lower() for t in out]

def content(name):
    c = [t for t in toks(name) if t not in STOP and len(t) > 1]
    return c or [t for t in toks(name) if len(t) > 1] or toks(name)

def op_name(op):
    for v in ("create", "move", "cancel"):
        if op.get(v):
            return op[v]
    return ""

def derive_node(ops):
    conts = {i: content(op_name(op)) for i, op in enumerate(ops)}
    freq = Counter(t for c in conts.values() for t in c)
    out = {}
    for i in conts:
        ranked = sorted(dict.fromkeys(conts[i]), key=lambda t: (freq[t], -len(t)))
        pick = ranked[:1]
        if ranked and freq[ranked[0]] > 1 and len(ranked) > 1:
            pick = ranked[:2]
        out[i] = pick or conts[i][:1]
    return out

def main():
    # load originals (for revert) keyed by (node_id, op_name)
    orig = {}
    for f in glob.glob(str(BK / "*.json")):
        n = json.load(open(f))
        for e in n.get("emails", []):
            for op in (e.get("answer") or {}).get("ops", []):
                orig[(n["id"], op_name(op))] = op.get("match")   # may be None

    # build in-memory nodes with derived match
    nodes = {}
    changed = 0
    for f in glob.glob(str(BK / "*.json")):
        n = json.load(open(f))
        for e in n.get("emails", []):
            ops = (e.get("answer") or {}).get("ops", [])
            if not ops:
                continue
            for i, m in derive_node(ops).items():
                if ops[i].get("match") != m:
                    changed += 1
                ops[i]["match"] = m
        nodes[n["id"]] = n

    def flush():
        for f in glob.glob(str(NODE_DIR / "*.json")):
            Path(f).unlink()
        for nid, n in nodes.items():
            (NODE_DIR / f"{nid}.json").write_text(json.dumps(n, indent=2) + "\n")

    reverted = []
    for _ in range(60):
        flush()
        try:
            load_corpus("corpus")
            break
        except CorpusError as ex:
            m = re.search(r"node '([^']+)': obligations '([^']+)' and '([^']+)'", str(ex))
            if not m:
                print("ABORT unexpected lint error:", str(ex)[:200]); sys.exit(1)
            nid, a, b = m.group(1), m.group(2), m.group(3)
            for e in nodes[nid].get("emails", []):
                for op in (e.get("answer") or {}).get("ops", []):
                    if op_name(op) in (a, b):
                        o = orig.get((nid, op_name(op)))
                        if o is None:
                            op.pop("match", None)
                        else:
                            op["match"] = o
                        reverted.append((nid, op_name(op)))
    else:
        print("ABORT: could not resolve all collisions"); sys.exit(1)

    corpus = load_corpus("corpus")
    plan = build_plan(corpus, start_date=date(2026,6,1), seed=42, n_days=730,
                      levers=Levers(daily_max=21))
    res = run(corpus, plan, oracle_model, store=Store(corpus))
    print(f"reloaded {len(corpus.emails)} emails; oracle {res.passed}/{res.total} = {res.score():.0%}")
    print(f"match sets changed: {changed}  |  reverted-to-original (ambiguous, flagged): {len(set(reverted))}")
    if res.score() < 1.0:
        print("ORACLE<100%:", [e for e,r in res.results.items() if not r.passed][:8]); sys.exit(1)

    print("\n=== SAMPLE new match keywords (op name -> match) ===")
    shown = 0
    for nid, n in sorted(nodes.items()):
        for e in n.get("emails", []):
            for op in (e.get("answer") or {}).get("ops", []):
                if shown < 26:
                    print(f"  [{nid[:20]:20}] {op_name(op)[:36]:36} -> {op.get('match', '(orig)')}")
                    shown += 1
    if reverted:
        print("\nFLAGGED ambiguous obligations (kept original phrase — need authoring):")
        for nid, nm in sorted(set(reverted)):
            print(f"  - [{nid}] {nm}")

if __name__ == "__main__":
    main()
