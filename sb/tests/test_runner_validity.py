"""
O-7: a day whose secretary tools were not working must not be scored as a model result.

The runner's day loop runs for real against real store and MCP servers on free
ports; only the `claude -p` call is replaced by a script that emits the CLI's
stream-json and, where a scenario needs it, acts on the store the way a model's
tool calls would. No model is involved.
"""
import json
import os
import shutil
import signal
import subprocess
import time
from datetime import date
from pathlib import Path

import httpx
import pytest

from sb import resolver
from sb.live import runner
from sb.regrade import ReferenceMismatch, check_reference
from sb.resolver import Context
from sb.schema import load_corpus
from sb.scheduler import Levers, build_plan

RAW_1C = Path("captures/baseline-sonnet-4-5/raw")
SERVED = "claude-sonnet-4-5-20250929"


def _stream(status, sid="sess-1"):
    """What `claude -p --output-format stream-json` prints, reduced to what the runner reads."""
    init = {"type": "system", "subtype": "init", "session_id": sid, "model": SERVED}
    if status != "no-secretary":
        init["mcp_servers"] = [{"name": "secretary", "status": status}]
    return "\n".join(json.dumps(e) for e in [
        init,
        {"type": "assistant", "message": {"id": "m1", "model": SERVED,
                                          "content": [{"type": "text", "text": "Done."}]}},
        {"type": "result", "is_error": False, "session_id": sid},
    ]) + "\n"


# --- the init-event detector --------------------------------------------------

def test_secretary_status_connected():
    assert runner._secretary_status(_stream("connected")) == "connected"


def test_secretary_status_failed():
    assert runner._secretary_status(_stream("failed")) == "failed"


def test_secretary_status_missing_server_or_init_event():
    assert runner._secretary_status(_stream("no-secretary")) == "absent"
    no_init = '{"type": "result", "is_error": false}\nnot json\n7\n'
    assert runner._secretary_status(no_init) is None
    assert runner._secretary_status("") is None


def test_secretary_status_on_the_1c_capture():
    """The 1c outage, read back from the raw streams: day 50 connected, days 51-57 failed."""
    status = {int(p.stem[4:]): runner._secretary_status(p.read_text())
              for p in sorted(RAW_1C.glob("day_*.jsonl"))}
    assert len(status) == 57
    assert {d for d, s in status.items() if s != "connected"} == set(range(51, 58))
    assert {status[d] for d in range(51, 58)} == {"failed"}


# --- the runner's day loop ----------------------------------------------------

@pytest.fixture
def harness(ports, tmp_path, monkeypatch):
    """Fast sleeps and health checks, a record of every MCP server started, and a
    scripted stand-in for `claude -p`. Set `h.script` to a list of callables."""
    real_run, real_sleep = subprocess.run, time.sleep
    real_healthy, real_start = runner._mcp_healthy, runner.start_mcp

    class H:
        script: list = []
        procs: list = []
        calls = 0
        out = tmp_path / "capture"

        def run(self, limit):
            runner.run("claude-sonnet-4-5", seed=42, start=date(2026, 6, 1), n_days=60,
                       limit=limit, driver="claude", levers=Levers(1, 5, 7),
                       out_dir=str(self.out))
            return (json.loads((self.out / "manifest.json").read_text()),
                    {int(p.stem): json.loads(p.read_text())
                     for p in (self.out / "days").glob("*.json")})

    h = H()

    def fake_run(cmd, *a, **kw):
        if cmd[0] != "claude":
            return real_run(cmd, *a, **kw)
        step = h.script[h.calls]
        h.calls += 1
        return step(cmd)

    def start(log_path):
        h.procs.append(real_start(log_path))
        return h.procs[-1]

    monkeypatch.setattr(runner.subprocess, "run", fake_run)
    monkeypatch.setattr(runner.time, "sleep", lambda s: real_sleep(min(s, 0.1)))
    monkeypatch.setattr(runner, "_mcp_healthy", lambda timeout=1.0: real_healthy(min(timeout, 1.0)))
    monkeypatch.setattr(runner, "start_mcp", start)
    yield h
    for p in h.procs:
        p.kill()
        p.wait()


def _turn(status):
    def step(cmd):
        return subprocess.CompletedProcess(cmd, 0, _stream(status), "")
    return step


def _acts_then(status):
    """The 'model' files a to-do for the first email of the day, then the turn ends."""
    def step(cmd):
        new = httpx.get(f"{runner.STORE_URL}/inbox/new", timeout=5).json()
        httpx.post(f"{runner.STORE_URL}/todos", timeout=5, json={
            "email_id": new[0]["email_id"], "title": "x", "due_date": "2026-06-05T17:00:00"})
        return subprocess.CompletedProcess(cmd, 0, _stream(status), "")
    return step


def _hangs_server_then(status, h):
    """The server wedges mid-turn: SIGSTOP leaves its port open but nothing answers."""
    def step(cmd):
        os.kill(h.procs[-1].pid, signal.SIGSTOP)
        return subprocess.CompletedProcess(cmd, 0, _stream(status), "")
    return step


def test_runner_restarts_the_secretary_server_and_retries(harness, monkeypatch):
    h = harness
    real_day = runner._Capture.day

    def day_then_hang(self, day_no, **fields):      # wedge the server between days 2 and 3
        real_day(self, day_no, **fields)
        if day_no == 2:
            os.kill(h.procs[-1].pid, signal.SIGSTOP)
    monkeypatch.setattr(runner._Capture, "day", day_then_hang)

    h.script = [_turn("failed"), _acts_then("connected"),          # day 1
                _hangs_server_then("connected", h), _turn("connected"),  # day 2
                _turn("connected")]                                 # day 3
    manifest, days = h.run(limit=9)                                 # days of 5, 2, 2 emails

    assert h.calls == 5
    assert manifest["mcp_restarts"] == 3 and len(h.procs) == 4
    assert manifest["errored"] == 0 and manifest["errored_days"] == []
    assert manifest["retried_days"] == [1, 2]

    d1 = days[1]
    assert d1["ok"] and d1["mcp_status"] == "connected"
    assert d1["attempts"][0] == {"attempt": 1, "ok": False, "mcp_status": "failed",
                                 "detail": "secretary MCP failed at turn start",
                                 "restarted_mcp": True}
    assert d1["attempts"][1] == {"attempt": 2, "ok": True, "mcp_status": "connected"}
    assert len(d1["day_new"]) == 1                     # the retry's to-do was graded
    assert (h.out / "raw" / "day_001.attempt_1.jsonl").exists()
    assert (h.out / "raw" / "day_001.jsonl").exists()

    assert days[2]["attempts"][0]["detail"] == "secretary MCP stopped answering during the turn"
    assert days[2]["attempts"][0]["restarted_mcp"] is True
    assert [a["ok"] for a in days[3]["attempts"]] == [True]     # restarted before the turn
    assert (h.out / "mcp_server.log").read_text().count("start_mcp") == 4
    assert all(p.poll() is not None for p in h.procs[:-1])


def test_day_is_errored_when_the_tools_never_connect(harness, monkeypatch):
    h = harness
    monkeypatch.setattr(runner, "ATTEMPTS", 3)
    h.script = [_turn("failed")] * 3
    manifest, days = h.run(limit=1)

    assert h.calls == 3
    d1 = days[1]
    assert d1["ok"] is False
    assert d1["error"] == "secretary MCP failed at turn start"
    assert [a["mcp_status"] for a in d1["attempts"]] == ["failed"] * 3
    assert manifest["errored"] == 1 and manifest["errored_days"] == [1]
    assert manifest["score_passed"] == 0 and manifest["score_total"] == 1
    verdicts = json.loads((h.out / "verdicts.json").read_text())
    assert list(verdicts.values()) == [None]


def test_a_turn_with_no_secretary_server_is_not_a_model_result(harness, monkeypatch):
    h = harness
    monkeypatch.setattr(runner, "ATTEMPTS", 2)
    h.script = [_turn("no-secretary"), _turn("connected")]
    manifest, days = h.run(limit=1)
    assert days[1]["attempts"][0]["detail"] == "secretary MCP absent at turn start"
    assert days[1]["ok"] and manifest["errored"] == 0


def test_a_failed_attempt_leaves_nothing_behind(harness):
    """Attempt 1 files a to-do and then times out. Without a rollback that to-do would sit
    in the retry's day delta, the duplicate O-7's original finding predicted."""
    h = harness

    def acts_then_times_out(cmd):
        _acts_then("connected")(cmd)
        raise subprocess.TimeoutExpired(cmd, 240, output=_stream("connected").encode())
    h.script = [acts_then_times_out, _turn("connected")]
    manifest, days = h.run(limit=1)

    d1 = days[1]
    assert d1["ok"] and [a["ok"] for a in d1["attempts"]] == [False, True]
    assert d1["attempts"][0]["detail"] == "turn timed out after 240s"
    assert d1["attempts"][0]["mcp_status"] == "connected"     # read from the partial stream
    assert "restarted_mcp" not in d1["attempts"][0]           # healthy server, not restarted
    assert d1["day_new"] == [] and d1["state"] == {"events": [], "todos": []}


def test_an_errored_day_leaves_nothing_behind(harness, monkeypatch):
    h = harness
    monkeypatch.setattr(runner, "ATTEMPTS", 2)
    h.script = [_acts_then("failed")] * 2
    manifest, days = h.run(limit=1)
    assert days[1]["ok"] is False
    assert json.loads((h.out / "state_final.json").read_text()) == {"events": [], "todos": []}


def test_capture_freezes_the_corpus_it_serves(harness, monkeypatch, tmp_path):
    """C-11: the capture carries the corpus it was served from, the store is posted exactly
    `_served_email` (same JSON as before C-11), and a --limit capture whose last day was cut
    short still passes the regrade reference check against a reformatted corpus."""
    h = harness
    posted = []
    real_post = runner.httpx.post

    def post(url, *a, **kw):
        if url.endswith("/inbox"):
            posted.append(kw["json"])
        return real_post(url, *a, **kw)
    monkeypatch.setattr(runner.httpx, "post", post)
    h.script = [_turn("connected")] * 3
    manifest, days = h.run(limit=8)                       # days of 5, 2 and 1 of 2 emails

    assert [len(days[d]["batch"]) for d in sorted(days)] == [5, 2, 1]
    assert (sorted(p.name for p in (h.out / "corpus" / "nodes").iterdir())
            == sorted(p.name for p in Path("corpus/nodes").glob("*.json")))
    assert runner._corpus_hash(str(h.out / "corpus")) == manifest["corpus_hash"]

    corpus = load_corpus("corpus")
    plan = build_plan(corpus, start_date=date(2026, 6, 1), seed=42, n_days=60, levers=Levers(1, 5, 7))
    expected = []
    for eid in (e for d in sorted(days) for e in days[d]["batch"]):
        email, sd = corpus.emails[eid], plan.serve_date[eid]
        expected.append({"email_id": eid, "sender": email.sender, "recipients": email.recipients,
                         "subject": email.subject,
                         "body": resolver.render_body(email.body, Context(sd, plan.anchors)).text,
                         "served_date": sd.isoformat()})
    assert [json.dumps(p) for p in posted] == [json.dumps(e) for e in expected]
    assert posted == [runner._served_email(corpus, plan, e["email_id"]) for e in expected]

    given = tmp_path / "given"
    shutil.copytree(h.out / "corpus", given)
    node = given / "nodes" / "Planning.json"
    node.write_text(json.dumps(json.loads(node.read_text()), indent=4))
    ref = check_reference(str(h.out), str(given))
    assert not ref.identical and ref.keys_changed == []
    manifest["limit"] = None                 # without --limit, a short last day is a mismatch
    (h.out / "manifest.json").write_text(json.dumps(manifest))
    with pytest.raises(ReferenceMismatch, match=r"day 3 \("):
        check_reference(str(h.out), str(given))


def test_offline_full_length_run(harness):
    """The offline stress test the approved rerun waits on: the real day loop over the
    1c plan (seed 42, 60 days -> 57 turns, 167 emails), with a scripted CLI that drives
    the real MCP server through the `mcp` client harder than the 1c model did (302
    secretary calls by day 50). The pre-O-7 server hung after ~64 KiB of stderr."""
    import asyncio

    from mcp import ClientSession
    from mcp.client.streamable_http import streamable_http_client

    h = harness
    calls = 0

    async def work_the_inbox():
        nonlocal calls
        async with streamable_http_client(runner.MCP_URL) as (r, w, _):
            async with ClientSession(r, w) as s:
                await s.initialize()
                await s.list_tools()
                new = (await s.call_tool("list_new_emails", {})).structuredContent["result"]
                calls += 1
                for m in new:
                    for name, args in [
                            ("get_email", {"email_id": m["email_id"]}),
                            ("search_inbox", {"query": m["subject"][:12]}),
                            ("get_email", {"email_id": m["email_id"]}),
                            ("create_todo", {"email_id": m["email_id"], "title": m["subject"],
                                             "due_date": "2026-09-01T17:00:00"})]:
                        res = await s.call_tool(name, args)
                        assert not res.isError, res
                        calls += 1

    def day(cmd):
        try:
            asyncio.run(asyncio.wait_for(work_the_inbox(), timeout=30))
        except Exception:
            # what the CLI would report on a turn that lost its tools
            return subprocess.CompletedProcess(cmd, 0, _stream("failed"), "")
        return subprocess.CompletedProcess(cmd, 0, _stream("connected"), "")

    h.script = [day] * 100                  # room for retries, so h.calls can catch one
    manifest, days = h.run(limit=None)

    assert len(days) == 57 and manifest["planned_emails"] == 167
    assert h.calls == 57, "a day needed a retry"
    assert calls == 57 + 4 * 167
    assert manifest["errored"] == 0 and manifest["mcp_restarts"] == 0
    assert {d["mcp_status"] for d in days.values()} == {"connected"}
    assert (h.out / "mcp_server.log").stat().st_size > 2 * 65536
