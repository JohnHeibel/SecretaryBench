"""
O-7: the secretary MCP server must keep answering for a whole run.

The 1c capture lost its secretary tools on days 51-57, most likely because
`start_mcp` piped the server's stderr and never read it. FastMCP logs every
request at INFO; once ~64 KiB sat unread in the pipe, the server blocked on a
log write and stopped answering, while the kernel kept accepting its TCP
connections. Every later `claude -p` turn saw the secretary server as "failed"
and the runner still recorded the days `ok`.

No model is involved. Real store and MCP servers run on free ports, driven by
the `mcp` Python client the way a `claude -p` turn drives them: a fresh session
per turn, then a day's worth of tool calls.
"""
import asyncio
import os
import socket
import subprocess
import sys
import time

import anyio
import httpx
import pytest
from mcp import ClientSession
from mcp.client.streamable_http import streamable_http_client

from sb.live import runner

PIPE_BUFFER = 65536          # macOS and Linux default pipe capacity
DAY_OF_CALLS = [("list_new_emails", {}), ("get_email", {"email_id": "n.e1"}),
                ("get_email", {"email_id": "n.e2"}),
                ("create_todo", {"email_id": "n.e1", "title": "t",
                                 "due_date": "2026-06-03T17:00:00"}),
                ("search_inbox", {"query": "policy"})]


@pytest.fixture
def store(ports):
    proc = runner.start_store()
    httpx.post(f"{runner.STORE_URL}/reset", timeout=10)
    httpx.post(f"{runner.STORE_URL}/day", timeout=10, json={"date": "2026-06-01"})
    for i in (1, 2):
        httpx.post(f"{runner.STORE_URL}/inbox", timeout=10, json={
            "email_id": f"n.e{i}", "sender": "s", "recipients": ["CEO"], "subject": "x",
            "body": "the meeting policy", "served_date": "2026-06-01"})
    yield proc
    proc.kill()
    proc.wait()


async def _turn(call_timeout: float) -> None:
    async with streamable_http_client(runner.MCP_URL) as (r, w, _):
        async with ClientSession(r, w) as s:
            with anyio.fail_after(call_timeout):
                await s.initialize()
                await s.list_tools()
            for name, args in DAY_OF_CALLS:
                with anyio.fail_after(call_timeout):
                    await s.call_tool(name, args)


def _drive(turns: int, call_timeout: float) -> int:
    """Run up to `turns` simulated turns; return how many completed before a hang."""
    async def go():
        for done in range(turns):
            try:
                await _turn(call_timeout)
            except Exception:       # a timed-out call surfaces as an ExceptionGroup
                return done
        return turns
    return asyncio.run(go())


def test_unread_stderr_pipe_blocks_the_server(store, ports):
    """The mechanism, with stderr piped and unread exactly as start_mcp did before O-7."""
    env = {**os.environ, "MCP_TRANSPORT": "streamable-http", "MCP_PORT": str(runner.MCP_PORT),
           "STORE_BASE_URL": runner.STORE_URL, "PYTHONPATH": str(runner.REPO)}
    proc = subprocess.Popen([sys.executable, "-m", "sb.live.mcp_app"], cwd=str(runner.REPO),
                            env=env, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
    try:
        for _ in range(60):
            if runner._mcp_healthy(timeout=2):
                break
            time.sleep(0.5)
        done = _drive(turns=300, call_timeout=3)
        assert done < 300, "server never blocked; the O-7 mechanism did not reproduce"

        assert proc.poll() is None, "server died rather than hung"
        with socket.socket() as s:
            s.settimeout(1)
            assert s.connect_ex(("127.0.0.1", runner.MCP_PORT)) == 0, \
                "a hung server should still accept TCP (why a port check is not a health check)"
        assert not runner._mcp_healthy(timeout=2)

        fd = proc.stderr.fileno()
        os.set_blocking(fd, False)
        backlog = 0
        while True:
            try:
                chunk = os.read(fd, PIPE_BUFFER)
            except BlockingIOError:
                break
            if not chunk:
                break
            backlog += len(chunk)
        assert backlog >= PIPE_BUFFER - 1024, f"only {backlog} bytes were waiting in the pipe"
        # With the pipe drained, the blocked log write completes and the server answers again.
        assert runner._mcp_healthy(timeout=5)
    finally:
        proc.kill()
        proc.wait()


def test_runner_mcp_server_survives_a_runs_worth_of_requests(store, tmp_path):
    """The fix: start_mcp logs to a file, so request volume cannot block the server.

    250 turns x 7 requests is several times what the 1c run sent before day 51
    (302 secretary tool calls over 50 turns)."""
    log = tmp_path / "mcp_server.log"
    proc = runner.start_mcp(log)
    try:
        assert _drive(turns=250, call_timeout=10) == 250
        assert runner._mcp_healthy()
        assert log.stat().st_size > 2 * PIPE_BUFFER, \
            "the stress test must write more than the pipe would have held"
    finally:
        proc.kill()
        proc.wait()


def test_start_mcp_reports_a_startup_crash(ports, tmp_path, monkeypatch):
    """E-3 still holds with stderr in a file: a server that dies on startup is
    reported with its traceback, not a generic "did not come up"."""
    real_popen = subprocess.Popen

    def crashing(cmd, **kw):
        return real_popen([sys.executable, "-c", "import sys; sys.exit('boom: bad mcp version')"],
                          **kw)
    monkeypatch.setattr(runner.subprocess, "Popen", crashing)
    with pytest.raises(RuntimeError, match="boom: bad mcp version"):
        runner.start_mcp(tmp_path / "mcp_server.log")
