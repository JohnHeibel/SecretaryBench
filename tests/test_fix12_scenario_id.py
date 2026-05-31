"""FIX-12: a write that 404s on an unregistered scenario_id is logged and
counted at the MCP server (so a garbled id is a visible diagnostic, not a
silent 0) while the error still surfaces to the model. No auto-correction.

Offline — _call is monkeypatched, no server needed."""
from __future__ import annotations

import pytest

from mcp_server import server


def test_bad_scenario_id_detected_and_counted(monkeypatch, capsys):
    def fake_call(method, path, **kw):
        if method == "POST":
            raise RuntimeError("HTTP 404 POST /todos/: {'detail': 'Scenario 12345 not found'}")
        if method == "GET" and path == "/scenarios/":
            return [{"scenario_id": 111}, {"scenario_id": 222}]
        return None

    monkeypatch.setattr(server, "_call", fake_call)
    before = server._bad_scenario_id_count

    with pytest.raises(RuntimeError):  # error still surfaces to the model
        server._write_checked("POST", "/todos/", 12345, json={})

    assert server._bad_scenario_id_count == before + 1
    err = capsys.readouterr().err
    assert "bad scenario_id" in err
    assert "seen=12345" in err
    assert "registered=[111, 222]" in err


def test_non_scenario_404_not_counted(monkeypatch, capsys):
    # A calendar-not-found 404 is NOT a scenario_id problem -> not counted.
    def fake_call(method, path, **kw):
        raise RuntimeError("HTTP 404 POST /calendars/x/events: {'detail': 'Calendar not found'}")

    monkeypatch.setattr(server, "_call", fake_call)
    before = server._bad_scenario_id_count

    with pytest.raises(RuntimeError):
        server._write_checked("POST", "/calendars/x/events", 111, json={})

    assert server._bad_scenario_id_count == before
    assert "bad scenario_id" not in capsys.readouterr().err


def test_successful_write_passes_through(monkeypatch):
    monkeypatch.setattr(server, "_call", lambda *a, **k: {"id": "ok"})
    assert server._write_checked("POST", "/todos/", 5, json={}) == {"id": "ok"}
