"""
authoring.server — the scenario-editor backend API.

Reuses the real engine (sb.resolver / sb.schema / sb.scheduler / sb.oracle) so
every preview the author sees is computed by the same code that runs the
benchmark. No database: corpus/nodes/*.json on disk is the source of truth.

    python -m authoring.server            # http://localhost:8099

Auth is stubbed: set AUTH_TOKEN to require an `X-Auth-Token` header (single
shared workspace); unset = open (dev default). Real per-user accounts slot in
here later without touching the editor.
"""
from __future__ import annotations

import os
from datetime import date, datetime
from pathlib import Path
from typing import Any, Optional

from fastapi import FastAPI, Header, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from authoring import chips, corpus_io
from sb import resolver
from sb.engine import Store, run
from sb.oracle import oracle_model
from sb.resolver import Context, Interval
from sb.schema import CorpusError, load_corpus
from sb.scheduler import InfeasibleSchedule, build_plan

CORPUS_DIR = Path(os.environ.get("CORPUS_DIR", "corpus"))
START = date(2026, 6, 1)
AUTH_TOKEN = os.environ.get("AUTH_TOKEN")  # None => open
WEB_DIST = Path(__file__).parent / "web" / "dist"

app = FastAPI(title="SecretaryBench Scenario Editor")
app.add_middleware(
    CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"],
)


def _auth(x_auth_token: Optional[str]) -> None:
    if AUTH_TOKEN and x_auth_token != AUTH_TOKEN:
        raise HTTPException(status_code=401, detail="bad or missing X-Auth-Token")


# --- request models --------------------------------------------------------

class ResolveReq(BaseModel):
    chip: Optional[dict] = None
    expr: Optional[str] = None
    serve: str                          # ISO date, the email's "today"
    anchors: dict[str, str] = {}        # name -> ISO date/datetime


class RenderReq(BaseModel):
    body_segments: Optional[list[dict]] = None
    body: Optional[str] = None
    serve: str
    anchors: dict[str, str] = {}


class ChipReq(BaseModel):
    chip: dict


class ParseReq(BaseModel):
    expr: str
    emit_as: Optional[str] = None


# --- helpers ---------------------------------------------------------------

def _parse_anchor_value(s: str):
    s = s.strip()
    try:
        return datetime.fromisoformat(s) if "T" in s else date.fromisoformat(s)
    except ValueError:
        raise HTTPException(status_code=422, detail=f"bad anchor value {s!r}")


def _ctx(serve: str, anchors: dict[str, str]) -> Context:
    try:
        serve_date = date.fromisoformat(serve)
    except ValueError:
        raise HTTPException(status_code=422, detail=f"bad serve date {serve!r}")
    return Context(serve=serve_date, anchors={k: _parse_anchor_value(v) for k, v in anchors.items()})


def _value_payload(v) -> dict:
    kind = resolver.value_kind(v)
    if isinstance(v, Interval):
        iso = v.start.isoformat()
    else:
        iso = v.isoformat()
    return {"ok": True, "kind": kind, "human": resolver.human(v), "iso": iso}


def _sample() -> dict:
    """A representative serve plan + anchor table, so '@signing+2w' previews resolve
    to a real date. Falls back to empty if the corpus is currently unsolvable."""
    try:
        corpus = load_corpus(CORPUS_DIR)
        plan = build_plan(corpus, start_date=START, seed=42, n_days=60)
    except (CorpusError, InfeasibleSchedule, Exception):  # noqa: BLE001
        return {"serve_date": {}, "anchors": {}, "ok": False}
    return {
        "ok": True,
        "serve_date": {eid: d.isoformat() for eid, d in plan.serve_date.items()},
        "anchors": {n: (v.start if isinstance(v, Interval) else v).isoformat()
                    for n, v in plan.anchors.items()},
    }


# ===========================================================================
# corpus
# ===========================================================================

@app.get("/api/corpus")
def get_corpus(x_auth_token: Optional[str] = Header(default=None)) -> dict:
    _auth(x_auth_token)
    graph = corpus_io.load_graph(CORPUS_DIR)
    graph["sample"] = _sample()
    graph["start"] = START.isoformat()
    return graph


@app.put("/api/thread/{thread_id}")
def put_thread(thread_id: str, thread: dict,
               x_auth_token: Optional[str] = Header(default=None)) -> dict:
    _auth(x_auth_token)
    thread.setdefault("id", thread_id)
    try:
        errors = corpus_io.write_node(CORPUS_DIR, thread)
    except chips.ChipError as exc:
        raise HTTPException(status_code=422, detail=f"chip error: {exc}")
    return {"errors": errors}


@app.delete("/api/thread/{thread_id}")
def delete_thread(thread_id: str, x_auth_token: Optional[str] = Header(default=None)) -> dict:
    _auth(x_auth_token)
    return {"errors": corpus_io.delete_node(CORPUS_DIR, thread_id)}


@app.post("/api/validate")
def validate(x_auth_token: Optional[str] = Header(default=None)) -> dict:
    _auth(x_auth_token)
    return {"errors": corpus_io.validate(CORPUS_DIR), "sample": _sample()}


# ===========================================================================
# live date engine
# ===========================================================================

@app.post("/api/resolve")
def resolve_chip(req: ResolveReq, x_auth_token: Optional[str] = Header(default=None)) -> dict:
    _auth(x_auth_token)
    try:
        expr = req.expr if req.expr is not None else chips.compile_chip(req.chip)
    except chips.ChipError as exc:
        return {"ok": False, "error": str(exc)}
    ctx = _ctx(req.serve, req.anchors)
    try:
        return _value_payload(resolver.resolve(expr, ctx))
    except resolver.ResolverError as exc:
        # most common: references an anchor we have no preview value for
        return {"ok": False, "error": str(exc)}


@app.post("/api/render")
def render_body(req: RenderReq, x_auth_token: Optional[str] = Header(default=None)) -> dict:
    _auth(x_auth_token)
    body = req.body if req.body is not None else corpus_io._segments_to_body(req.body_segments or [])
    ctx = _ctx(req.serve, req.anchors)
    try:
        out = resolver.render_body(body, ctx)
    except resolver.ResolverError as exc:
        return {"ok": False, "error": str(exc), "body": body}
    return {"ok": True, "text": out.text,
            "emissions": {n: resolver.human(v) for n, v in out.emissions.items()}}


@app.post("/api/chip/compile")
def chip_compile(req: ChipReq, x_auth_token: Optional[str] = Header(default=None)) -> dict:
    _auth(x_auth_token)
    try:
        return {"ok": True, "expr": chips.compile_chip(req.chip),
                "token": chips.compile_body_token(req.chip)}
    except chips.ChipError as exc:
        return {"ok": False, "error": str(exc)}


@app.post("/api/chip/parse")
def chip_parse(req: ParseReq, x_auth_token: Optional[str] = Header(default=None)) -> dict:
    _auth(x_auth_token)
    return {"ok": True, "chip": chips.parse_token(req.expr, emit_as=req.emit_as)}


# ===========================================================================
# oracle satisfiability — "can a perfect secretary score 100% on this?"
# ===========================================================================

@app.post("/api/oracle")
def oracle_check(x_auth_token: Optional[str] = Header(default=None)) -> dict:
    _auth(x_auth_token)
    try:
        corpus = load_corpus(CORPUS_DIR)
    except CorpusError as exc:
        return {"ok": False, "error": str(exc), "results": {}}
    try:
        plan = build_plan(corpus, start_date=START, seed=42, n_days=60)
    except InfeasibleSchedule as exc:
        return {"ok": False, "error": f"infeasible schedule: {exc}", "results": {}}

    store = Store(corpus)
    res = run(corpus, plan, oracle_model, store=store)
    results = {eid: {"passed": r.passed, "headline": r.headline}
               for eid, r in res.results.items()}
    return {"ok": True, "score": res.score(), "passed": res.passed,
            "total": res.total, "results": results}


@app.get("/api/health")
def health() -> dict:
    return {"ok": True}


# --- static frontend (built) ----------------------------------------------
# Mounted last so /api/* wins. In dev the Vite server proxies here instead.

if WEB_DIST.exists():
    app.mount("/assets", StaticFiles(directory=WEB_DIST / "assets"), name="assets")

    @app.get("/{full_path:path}")
    def spa(full_path: str):
        target = WEB_DIST / full_path
        if full_path and target.is_file():
            return FileResponse(target)
        return FileResponse(WEB_DIST / "index.html")


def main() -> None:
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=int(os.environ.get("PORT", "8099")))


if __name__ == "__main__":
    main()
