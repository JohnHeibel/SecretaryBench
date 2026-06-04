"""POST /api/lint — load + lint a whole corpus IN MEMORY with the real
sb.schema.build_corpus. This is the gate: the UI blocks save unless this returns
ok. It enforces exactly what the benchmark enforces — acyclic DAG, every anchor
emitted by a transitive ancestor, every token/predicate parses, and the
"answer references @anchor => must carry a date edge" rule.

Request:  { "nodes": [ { "id": "...", "cast": {...}, "emails": [...] }, ... ] }
Response (ok):   { "ok": true, "summary": {...}, "emission_map": {...},
                   "emails": [ {id, node, depends_on, emits, anchor_refs} ],
                   "order": [email_id, ...] }
Response (bad):  { "ok": false, "error": "email 'x' references undefined anchor @y" }
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "_lib"))

from http.server import BaseHTTPRequestHandler

import apihelp
from sb import schema


def _lint(raw_nodes: list) -> dict:
    corpus = schema.build_corpus(raw_nodes)        # raises CorpusError on any problem
    emails = [
        {
            "id": e.id,
            "node": e.node,
            "depends_on": [{"email": d.email, "type": d.type} for d in e.depends_on],
            "emits": sorted(e.emits.keys()),
            "anchor_refs": sorted(e.anchor_refs),
        }
        for e in corpus.emails.values()
    ]
    return {
        "ok": True,
        "summary": {
            "nodes": len(corpus.nodes),
            "emails": len(corpus.emails),
            "anchors": len(corpus.emission_map),
        },
        "emission_map": dict(corpus.emission_map),
        "emails": emails,
        "order": corpus.topo_order(),
    }


class handler(BaseHTTPRequestHandler):
    def do_POST(self):                                  # noqa: N802 (Vercel contract)
        try:
            body = apihelp.read_json(self)
            nodes = body.get("nodes")
            if not isinstance(nodes, list):
                return apihelp.send_json(self, 400, {"ok": False, "error": "expected 'nodes': []"})
            apihelp.send_json(self, 200, _lint(nodes))
        except schema.CorpusError as exc:
            apihelp.send_json(self, 200, {"ok": False, "error": str(exc)})
        except Exception as exc:
            apihelp.send_json(self, 200, {"ok": False, "error": f"{type(exc).__name__}: {exc}"})
