import socket

import pytest

from sb.live import runner


def _free_tcp_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


@pytest.fixture
def ports(monkeypatch):
    """Point the runner at free ports, so a test never touches a real run's servers
    (`runner._free_port` SIGKILLs whatever holds the port)."""
    store_port, mcp_port = _free_tcp_port(), _free_tcp_port()
    monkeypatch.setattr(runner, "STORE_PORT", store_port)
    monkeypatch.setattr(runner, "STORE_URL", f"http://localhost:{store_port}")
    monkeypatch.setattr(runner, "MCP_PORT", mcp_port)
    monkeypatch.setattr(runner, "MCP_URL", f"http://127.0.0.1:{mcp_port}/mcp")
    return store_port, mcp_port
