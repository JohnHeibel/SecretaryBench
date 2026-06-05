"""POST /api/resolve — evaluate ONE grammar expression against a serve date and
optional anchor table, using the real sb.resolver. Powers the block editor's live
"this token resolves to ..." preview and the DAG canvas date hints.

Request:  { "expr": "@signing+2w", "serve_date": "2026-06-01",
            "anchors": { "signing": "2026-06-08" } }
Response: { "ok": true, "kind": "date", "iso": "2026-06-22",
            "human": "Monday, June 22nd, 2026" }
          { "ok": true, "kind": "timeinterval", "iso": "2026-06-04T14:00:00/2026-06-04T15:00:00",
            "human": "Thursday, June 4th, 2026, 2–3 PM" }   for a @HH:MM-HH:MM expr
          { "ok": false, "error": "..." }   on a bad expression
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "_lib"))

from http.server import BaseHTTPRequestHandler

import apihelp
from sb import resolver


def _evaluate(expr: str, serve_date: str, anchors: dict) -> dict:
    ctx = resolver.Context(
        serve=apihelp.parse_date(serve_date),
        anchors={name: apihelp.parse_date(v) for name, v in (anchors or {}).items()},
    )
    value = resolver.resolve(expr, ctx)
    kind = resolver.value_kind(value)
    if isinstance(value, resolver.Interval):                     # whole-day span [start..end]
        iso = f"{value.start.isoformat()}..{value.end.isoformat()}"
    elif isinstance(value, resolver.TimeInterval):               # within-day clock span [start/end]
        iso = f"{value.start.isoformat()}/{value.end.isoformat()}"
    else:                                                        # date or datetime — both isoformat()
        iso = value.isoformat()
    return {"ok": True, "kind": kind, "iso": iso, "human": resolver.human(value)}


class handler(BaseHTTPRequestHandler):
    def do_POST(self):                                  # noqa: N802 (Vercel contract)
        try:
            body = apihelp.read_json(self)
            expr = (body.get("expr") or "").strip()
            if not expr:
                return apihelp.send_json(self, 400, {"ok": False, "error": "missing 'expr'"})
            serve = body.get("serve_date") or "2026-06-01"
            result = _evaluate(expr, serve, body.get("anchors") or {})
            apihelp.send_json(self, 200, result)
        except resolver.ResolverError as exc:
            apihelp.send_json(self, 200, {"ok": False, "error": str(exc)})
        except Exception as exc:                        # malformed date, bad JSON, etc.
            apihelp.send_json(self, 200, {"ok": False, "error": f"{type(exc).__name__}: {exc}"})
