#!/usr/bin/env python3
"""Local dev validator. `next dev` cannot run the Vercel Python functions, so this
serves the SAME /api/resolve, /api/lint and /api/oracle logic (reusing
api/resolve.py, api/lint.py and api/oracle.py — no second implementation) on a side
port with CORS, for the browser to call during development. On Vercel the real
functions serve these paths instead.

    python3 scripts/validator_server.py [port]      # default 8090
"""
from __future__ import annotations

import importlib.util
import json
import os
import sys
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HERE = os.path.dirname(os.path.abspath(__file__))
API = os.path.join(HERE, "..", "api")
sys.path.insert(0, os.path.join(API, "_lib"))


def _load(name: str):
    spec = importlib.util.spec_from_file_location(name, os.path.join(API, f"{name}.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


resolve_mod = _load("resolve")
lint_mod = _load("lint")
oracle_mod = _load("oracle")
from sb import resolver, schema   # noqa: E402  (path set above)
from sb.scheduler import InfeasibleSchedule   # noqa: E402


class Handler(BaseHTTPRequestHandler):
    def _cors(self):
        self.send_header("access-control-allow-origin", "*")
        self.send_header("access-control-allow-headers", "content-type")
        self.send_header("access-control-allow-methods", "POST, OPTIONS")

    def do_OPTIONS(self):  # noqa: N802
        self.send_response(204)
        self._cors()
        self.end_headers()

    def _send(self, status: int, payload: dict):
        body = json.dumps(payload).encode()
        self.send_response(status)
        self.send_header("content-type", "application/json")
        self._cors()
        self.send_header("content-length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):  # noqa: N802
        length = int(self.headers.get("content-length") or 0)
        body = json.loads(self.rfile.read(length) or b"{}") if length else {}
        try:
            if self.path.rstrip("/") == "/api/resolve":
                expr = (body.get("expr") or "").strip()
                if not expr:
                    return self._send(400, {"ok": False, "error": "missing 'expr'"})
                return self._send(200, resolve_mod._evaluate(expr, body.get("serve_date") or "2026-06-01", body.get("anchors") or {}))
            if self.path.rstrip("/") == "/api/lint":
                nodes = body.get("nodes")
                if not isinstance(nodes, list):
                    return self._send(400, {"ok": False, "error": "expected 'nodes': []"})
                return self._send(200, lint_mod._lint(nodes))
            if self.path.rstrip("/") == "/api/oracle":
                nodes = body.get("nodes")
                if not isinstance(nodes, list):
                    return self._send(400, {"ok": False, "error": "expected 'nodes': []"})
                return self._send(200, oracle_mod._oracle(nodes))
            self._send(404, {"ok": False, "error": f"no route {self.path}"})
        except (resolver.ResolverError, schema.CorpusError, InfeasibleSchedule) as exc:
            self._send(200, {"ok": False, "error": str(exc)})
        except Exception as exc:
            self._send(200, {"ok": False, "error": f"{type(exc).__name__}: {exc}"})

    def log_message(self, *_):  # quiet
        pass


if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8090
    print(f"validator on http://localhost:{port}  (/api/resolve, /api/lint, /api/oracle)")
    ThreadingHTTPServer(("127.0.0.1", port), Handler).serve_forever()
