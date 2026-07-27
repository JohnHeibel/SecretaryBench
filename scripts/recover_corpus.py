"""
Local, throwaway corpus recovery. Only ever REMOVE what is genuinely broken; never
invent an answer. Grading-lossless (main grades whole-day only). Every exclusion is
reported so the team can fix the source.

Method:
  1. Time-strip vestigial @HH:MM suffixes; drop blank ops (abandoned stubs).
  2. CLASSIFY each scenario IN ISOLATION: build + schedule-alone(730d) + oracle-alone.
     A scenario that fails alone is a real authoring bug -> exclude + reason.
     (This separates true breakage from capacity effects of interleaving.)
  3. From GOOD scenarios, resolve global anchor collisions (drop the stub = fewer ops).
  4. Assemble combined; pick scheduler levers (raise daily_max) so capacity never
     strands an individually-feasible scenario.
  5. GATE: combined oracle MUST be 100%, else report the offending email and abort.
"""
import json, os, re, sys
from pathlib import Path
from datetime import date
import httpx
from sb.schema import build_corpus, load_corpus
from sb.scheduler import build_plan, InfeasibleSchedule, Levers
from sb.engine import Store, run
from sb.oracle import oracle_model

PROD_URL = os.environ.get("SB_WEBAPP_URL", "https://secretarybench.vercel.app").rstrip("/") + "/api/nodes"
REPO = Path(__file__).resolve().parents[1]
NODE_DIR = REPO / "corpus" / "nodes"
START, SEED, N_DAYS = date(2026, 6, 1), 42, 730
_TIME = re.compile(r"\s*@\d{1,2}:\d{2}(?:-\d{1,2}:\d{2})?")

def fetch_prod():
    data = httpx.get(PROD_URL, timeout=30).json()
    return data.get("nodes", data) if isinstance(data, dict) else data

def strip_times(o, c):
    if isinstance(o, str): c[0]+=len(_TIME.findall(o)); return _TIME.sub("", o)
    if isinstance(o, list): return [strip_times(x, c) for x in o]
    if isinstance(o, dict): return {k: strip_times(v, c) for k, v in o.items()}
    return o

def is_blank_op(op): return any(v in op and not op[v] for v in ("create","move","cancel"))
def gops(n):
    try: c=build_corpus([n]); return sum(len(e.answer.ops) for e in c.emails.values())
    except Exception: return -1
def e2n(eid, ids):
    cands=[i for i in ids if eid==i or eid.startswith(i+".")]; return max(cands,key=len) if cands else None

def classify_alone(n):
    """Return (ok, reason)."""
    try: c = build_corpus([n])
    except Exception as ex: return False, f"parse: {str(ex)[:80]}"
    if not any(e.answer.ops for e in c.emails.values()):
        return True, "no-op (pure distractor)"     # valid, trivially schedulable
    try: plan = build_plan(c, start_date=START, seed=SEED, n_days=N_DAYS,
                           levers=Levers(daily_max=50))
    except InfeasibleSchedule as ex: return False, f"schedule: {str(ex)[:80]}"
    store=Store(c); res=run(c, plan, oracle_model, store=store)
    if res.score() < 1.0:
        bad=[eid for eid,r in res.results.items() if not r.passed]
        return False, f"oracle {res.passed}/{res.total}: unsatisfiable answer in {bad}"
    return True, "ok"

def main():
    nodes = fetch_prod(); tc=[0]; dropped=0
    xf={}
    for n in nodes:
        n2=strip_times(json.loads(json.dumps(n)), tc)
        for e in n2.get("emails",[]):
            a=e.get("answer") or {}
            if "ops" in a:
                b=len(a["ops"]); a["ops"]=[op for op in a["ops"] if not is_blank_op(op)]; dropped+=b-len(a["ops"])
        xf[n2["id"]]=n2

    # (2) classify alone
    good, excluded = {}, []
    print("PER-SCENARIO CLASSIFICATION (in isolation):")
    for nid,n in xf.items():
        ok,reason = classify_alone(n)
        ne=len(n.get("emails",[]))
        tag="GOOD" if ok else "BROKEN"
        if ne or not ok: print(f"  {tag:6s} {nid:42s} emails={ne:2d}  {reason}")
        (good if ok else None)
        if ok: good[nid]=n
        else: excluded.append((nid, reason))

    # (3) resolve global anchor collisions among GOOD
    for _ in range(50):
        try: build_corpus(list(good.values())); break
        except Exception as ex:
            m=re.search(r"emitted by both '([^']+)' and '([^']+)'", str(ex))
            if not m: print(f"ABORT combined build: {str(ex)[:160]}"); sys.exit(1)
            a,b=e2n(m.group(1),good), e2n(m.group(2),good)
            victim=min([a,b], key=lambda i: gops(good[i]))
            keep=b if victim==a else a
            excluded.append((victim, f"dup global anchor with '{keep}' (kept larger)"))
            good.pop(victim)
    corpus=build_corpus(list(good.values()))

    # (4) combined schedule — escalate daily_max so capacity never strands a good scenario
    plan=None
    for dmax in (5,8,12,20,40,80):
        try:
            plan=build_plan(corpus, start_date=START, seed=SEED, n_days=N_DAYS,
                            levers=Levers(daily_max=dmax)); chosen=dmax; break
        except InfeasibleSchedule as ex: last=str(ex)
    if plan is None:
        print(f"ABORT combined schedule even at daily_max=80: {last[:160]}"); sys.exit(1)

    # (5) oracle gate combined
    store=Store(corpus); res=run(corpus, plan, oracle_model, store=store)
    ndays=sum(1 for b in plan.per_day if b); last=max((i for i,b in enumerate(plan.per_day) if b),default=0)
    real=[n for n in good.values() if n.get("emails")]
    print("="*72)
    print(f"time tokens stripped={tc[0]}  blank ops dropped={dropped}")
    print(f"INCLUDED {len(good)} scenarios ({len(real)} with emails); EXCLUDED {len(excluded)}")
    print(f"emails={len(corpus.emails)} graded_ops={sum(len(e.answer.ops) for e in corpus.emails.values())}")
    print(f"scheduler: daily_max={chosen}, {ndays} non-empty days, last day idx {last} (ceiling {N_DAYS})")
    print(f"ORACLE combined: {res.passed}/{res.total} = {res.score():.0%}")
    print("-"*72+"\nINCLUDED (emails/ops):")
    for n in sorted(real, key=lambda n:-len(n["emails"])):
        print(f"  + {n['id']:44s} {len(n['emails']):2d} / {gops(n)}")
    emp=[n['id'] for n in good.values() if not n.get('emails')]
    if emp: print(f"  (+ {len(emp)} empty stubs: {', '.join(emp)})")
    print("EXCLUDED (needs source fix):")
    for nid,why in excluded: print(f"  - {nid:44s} {why}")

    if res.score()<1.0:
        bad=[eid for eid,r in res.results.items() if not r.passed]
        print(f"\nABORT: combined oracle <100%, offending {bad} — NOT writing."); sys.exit(1)

    for f in NODE_DIR.glob("*.json"): f.unlink()
    NODE_DIR.mkdir(parents=True, exist_ok=True)
    for n in good.values(): (NODE_DIR/f"{n['id']}.json").write_text(json.dumps(n,indent=2)+"\n")
    print(f"\nWROTE {len(good)} node files -> corpus/nodes/; reload:", len(load_corpus('corpus').emails), "emails")
    print(f"RECOMMENDED RUN LEVERS: daily_max={chosen}, n_days>={last+1}")

if __name__=="__main__": main()
