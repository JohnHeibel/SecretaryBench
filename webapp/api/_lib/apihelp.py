"""Tiny shared helpers for the Vercel Python serverless functions. Puts the
vendored `sb` package on the path and gives a uniform JSON request/response
shape. (Lives in api/_lib so Vercel does not treat it as an endpoint.)"""
from __future__ import annotations

import json
import os
import sys
from datetime import date

_LIB = os.path.dirname(__file__)
if _LIB not in sys.path:
    sys.path.insert(0, _LIB)          # so `from sb import ...` finds api/_lib/sb


# Cap request bodies so an oversized payload can't be read wholesale into memory
# (defense-in-depth: these functions are passcode-gated, but the solver is CPU-heavy).
_MAX_BODY = 2 * 1024 * 1024  # 2 MB — a corpus of authored emails is far smaller


class RequestTooLarge(ValueError):
    pass


def read_json(handler) -> dict:
    length = int(handler.headers.get("content-length") or 0)
    if length > _MAX_BODY:
        raise RequestTooLarge(f"request body {length} bytes exceeds {_MAX_BODY} limit")
    if not length:
        return {}
    raw = handler.rfile.read(length)
    return json.loads(raw or b"{}")


def send_json(handler, status: int, payload: dict) -> None:
    body = json.dumps(payload).encode()
    handler.send_response(status)
    handler.send_header("content-type", "application/json")
    handler.send_header("content-length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def parse_date(s: str) -> date:
    return date.fromisoformat(s)
