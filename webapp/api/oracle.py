"""POST /api/oracle — SATISFIABILITY check for a whole corpus IN MEMORY. The lint
gate proves the corpus is well-FORMED (acyclic DAG, anchors reachable, tokens
parse); this proves it is well-ANSWERED. It runs the real reference solver
(sb.oracle.oracle_model) — a perfect secretary that acts straight from the answer
key — through the real serve/grade loop. If the oracle can't score 1.0, some
answer key is unsatisfiable: the linter would happily pass it, but no model could.

corpus = schema.build_corpus(nodes) → plan = scheduler.build_plan(seed=42,
start_date 2026-06-01, n_days=...) → engine.run(.., oracle_model). ok = score==1.0.

The serve window scales with the corpus: the scheduler serves a few emails a day, so
a fixed 120-day window strands anything past ~350 emails. Default n_days to comfortably
cover the corpus (>= emails + slack), and let a caller override via "n_days".

Request:  { "nodes": [ {...}, ... ], "n_days"?: 500 }
Response (ok):   { "ok": true,  "score": 1.0, "passed": N, "total": N, "failures": [] }
Response (bad):  { "ok": false, "score": 0.8, "passed": 4, "total": 5,
                   "failures": ["node-1.e2", ...] }
Response (err):  { "ok": false, "error": "..." }   on a malformed/infeasible corpus
"""
from __future__ import annotations

import os
import sys
from datetime import date

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "_lib"))

from http.server import BaseHTTPRequestHandler

import apihelp
from sb import engine, schema
from sb.oracle import oracle_model
from sb.scheduler import InfeasibleSchedule, build_plan


def _oracle(raw_nodes: list, n_days: int | None = None) -> dict:
    corpus = schema.build_corpus(raw_nodes)            # raises CorpusError
    # ~daily_max serves/day means feasibility needs days on the order of the email count;
    # default generously (the scheduler is cheap now) and honor an explicit override.
    days = n_days if n_days else max(120, len(corpus.emails) + 120)
    plan = build_plan(corpus, start_date=date(2026, 6, 1), seed=42, n_days=days)  # raises InfeasibleSchedule
    res = engine.run(corpus, plan, oracle_model)
    failures = sorted(eid for eid, r in res.results.items() if not r.passed)
    return {
        "ok": res.score() == 1.0,
        "score": res.score(),
        "passed": res.passed,
        "total": res.total,
        "failures": failures,
    }


class handler(BaseHTTPRequestHandler):
    def do_POST(self):                                  # noqa: N802 (Vercel contract)
        try:
            body = apihelp.read_json(self)
            nodes = body.get("nodes")
            if not isinstance(nodes, list):
                return apihelp.send_json(self, 400, {"ok": False, "error": "expected 'nodes': []"})
            n_days = body.get("n_days")
            apihelp.send_json(self, 200, _oracle(nodes, n_days if isinstance(n_days, int) else None))
        except (schema.CorpusError, InfeasibleSchedule) as exc:
            apihelp.send_json(self, 200, {"ok": False, "error": str(exc)})
        except Exception as exc:
            apihelp.send_json(self, 200, {"ok": False, "error": f"{type(exc).__name__}: {exc}"})
