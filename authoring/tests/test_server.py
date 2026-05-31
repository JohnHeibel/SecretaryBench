"""API smoke tests against the real corpus (read-only; no disk writes)."""
from __future__ import annotations

from fastapi.testclient import TestClient

from authoring.server import app

c = TestClient(app)


def test_corpus_loads_clean():
    g = c.get("/api/corpus").json()
    assert g["errors"] == []
    assert {t["id"] for t in g["threads"]} == {"acme-client", "henderson", "hr-policy"}
    assert g["sample"]["ok"] is True
    assert g["sample"]["anchors"].get("signing")


def test_resolve_downstream_anchor_chip():
    g = c.get("/api/corpus").json()
    serve = g["sample"]["serve_date"]["henderson.kickoff"]
    r = c.post("/api/resolve", json={
        "chip": {"base": {"kind": "anchor", "name": "signing"},
                 "offset": {"amount": 2, "unit": "weeks"},
                 "time": {"hour": 9, "minute": 0}},
        "serve": serve, "anchors": g["sample"]["anchors"]}).json()
    assert r["ok"] and r["kind"] == "datetime"
    assert r["iso"].endswith("T09:00:00")


def test_resolve_unknown_anchor_is_graceful():
    r = c.post("/api/resolve", json={"expr": "@nope+1w", "serve": "2026-06-01", "anchors": {}}).json()
    assert r["ok"] is False and "nope" in r["error"]


def test_render_reports_emissions():
    r = c.post("/api/render", json={"body": "Locked for {!signing = serve+5d}.",
                                    "serve": "2026-06-01", "anchors": {}}).json()
    assert r["ok"] and "signing" in r["emissions"]


def test_chip_compile_and_parse_round_trip():
    cc = c.post("/api/chip/compile", json={
        "chip": {"base": {"kind": "next_weekday", "weekday": "THU"},
                 "offset": None, "time": {"hour": 14, "minute": 0}}}).json()
    assert cc["expr"] == "next:THU@14:00" and cc["token"] == "{next:THU@14:00}"
    pp = c.post("/api/chip/parse", json={"expr": cc["expr"]}).json()
    assert pp["chip"]["base"]["kind"] == "next_weekday"


def test_oracle_is_fully_satisfiable():
    o = c.post("/api/oracle").json()
    assert o["ok"] and o["score"] == 1.0 and o["passed"] == o["total"]
