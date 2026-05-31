"""
sb.live.runner — drive a REAL model over the corpus and grade it.

Architecture for one run:
    [this process] starts uvicorn store_app (persistent state)
        for each email in plan order:
            POST the rendered email to the store inbox
            call `claude -p` (ONE continuous session, --resume) -> model acts via MCP
            snapshot /state -> grade the email against its node state
        tear down

    python -m sb.live.runner --model claude-haiku-4-5 --seed 42
"""
from __future__ import annotations

import argparse
import json
import os
import signal
import socket
import subprocess
import sys
import time
from datetime import date, datetime
from pathlib import Path
from typing import Optional

import httpx

from sb import resolver
from sb.grader import NodeState, Obj, TurnDelta, grade_email
from sb.resolver import Context
from sb.schema import Corpus, load_corpus
from sb.scheduler import build_plan

REPO = Path(__file__).resolve().parents[2]
STORE_PORT = int(os.environ.get("STORE_PORT", "8077"))
STORE_URL = f"http://localhost:{STORE_PORT}"
MCP_PORT = int(os.environ.get("MCP_PORT", "8078"))
MCP_URL = f"http://127.0.0.1:{MCP_PORT}/mcp"

# --- ANSI palette (disabled if stdout isn't a tty or NO_COLOR is set) -------
_COLOR = (sys.stdout.isatty() or os.environ.get("FORCE_COLOR")) and not os.environ.get("NO_COLOR")
def _c(code: str) -> str:
    return code if _COLOR else ""
RESET, BOLD, DIM = _c("\033[0m"), _c("\033[1m"), _c("\033[2m")
RED, GREEN, YELLOW = _c("\033[31m"), _c("\033[32m"), _c("\033[33m")
CYAN, GREY, BLUE = _c("\033[36m"), _c("\033[90m"), _c("\033[34m")


def _print_email(i: int, eid: str, served: str, res, tools: list[str], said: str) -> None:
    ok = res.passed
    badge = f"{GREEN}{BOLD} ✓ PASS {RESET}" if ok else f"{RED}{BOLD} ✗ FAIL {RESET}"
    print(f"\n{badge} {BOLD}[{i}] {eid}{RESET}  {GREY}· served {served}{RESET}")
    for d in res.details:
        tick = f"{GREEN}✓{RESET}" if d["passed"] else f"{RED}✗{RESET}"
        rc = GREEN if d["passed"] else RED
        print(f"   {tick} {DIM}expected{RESET}  {CYAN}{d['expected']}{RESET}")
        print(f"     {DIM}actual  {RESET}  {d['actual']}")
        print(f"     {DIM}why     {RESET}  {rc}{d['reason']}{RESET}")
    tool_str = ", ".join(tools) if tools else "(none)"
    searched = f"  {CYAN}🔍 searched inbox{RESET}" if any("inbox" in t or t == "get_email" for t in tools) else ""
    print(f"     {DIM}tools   {RESET}  {GREY}{tool_str}{RESET}{searched}")
    if not ok and said:
        snippet = said[:200].replace("\n", " ")
        print(f"     {DIM}model said{RESET}  {YELLOW}“{snippet}…”{RESET}")

SYSTEM_PROMPT = """You are an executive assistant managing one person's calendar and to-dos from their email inbox. You receive emails ONE AT A TIME over many simulated days, and you keep ONE continuous memory across all of them.

For each email, decide the correct action(s) and call the matching tool — or do nothing. Do not narrate; stop after the tool call(s).

TOOLS:
- create_event / update_event / delete_event — the calendar.
- create_todo / update_todo / delete_todo — tasks that have a deadline.
- search_inbox / get_email — look back at earlier emails. USE THIS whenever an email refers to a fact, date, policy, person, or decision you don't fully remember (e.g. "two weeks after the signing", "per the new meeting policy"). The information was usually set in an earlier email.
- (no tool) — for FYIs, newsletters, confirmations, or anything needing no calendar/to-do change. Over-acting is the most common mistake; when in doubt, do nothing.

EVERY create_event and create_todo REQUIRES the email_id of the email you are acting on. It is given in the header as "Email-Id: <id>". Copy it exactly.

DATES: each email header gives "Today: <date>". Compute all relative dates from it. Use ISO 8601 in tool calls. Defaults unless the email says otherwise: events 9am, 1 hour long; to-dos due 5pm.

RESCHEDULING: to move an event, update_event it (delete-and-recreate is also fine). To cancel, delete it. Never leave duplicates."""


def _mcp_config() -> str:
    # Attach to the ONE persistent server (see start_mcp) over HTTP, instead of
    # cold-spawning a stdio server per turn — which raced --resume turns on fast models.
    return json.dumps({"mcpServers": {"secretary": {"type": "http", "url": MCP_URL}}})


def _parse_iso(s: str) -> Optional[datetime]:
    if not s:
        return None
    try:
        return datetime.fromisoformat(s.replace("Z", "").split("+")[0])
    except ValueError:
        try:
            return datetime.fromisoformat(s.split("T")[0])
        except ValueError:
            return None


def _obj_from(record: dict, kind: str) -> Optional[Obj]:
    when = _parse_iso(record.get("start") or record.get("due_date") or "")
    if when is None:
        return None
    return Obj(kind=kind, title=record.get("title", ""), when=when,
               email_id=record.get("email_id", ""), description=record.get("description", ""),
               end=_parse_iso(record.get("end", "")) if kind == "event" else None)


def _node_state(corpus: Corpus, state: dict, node: str, sid_filter: set[str]) -> NodeState:
    def keep(rec, kind):
        o = _obj_from(rec, kind)
        if o is None:
            return None
        e = corpus.emails.get(o.email_id)
        return o if (e and e.node == node) else None
    ev = [o for r in state["events"] if (o := keep(r, "event"))]
    td = [o for r in state["todos"] if (o := keep(r, "todo"))]
    return NodeState(events=ev, todos=td)


def _turn_delta(corpus: Corpus, state: dict, new_ids: set[str]) -> TurnDelta:
    ev = [o for r in state["events"] if r["id"] in new_ids and (o := _obj_from(r, "event"))]
    td = [o for r in state["todos"] if r["id"] in new_ids and (o := _obj_from(r, "todo"))]
    return TurnDelta(events=ev, todos=td)


def _all_ids(state: dict) -> set[str]:
    return {r["id"] for r in state["events"]} | {r["id"] for r in state["todos"]}


def _parse_stream(stdout: str) -> tuple[Optional[str], list[str], bool, str]:
    """Return (session_id, tool_names_used, is_error, assistant_text)."""
    session_id, tools, is_error, texts = None, [], False, []
    seen = {}
    for line in stdout.splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            ev = json.loads(line)
        except json.JSONDecodeError:
            continue
        t = ev.get("type")
        if t == "system" and ev.get("subtype") == "init":
            session_id = ev.get("session_id")
        elif t == "assistant":
            msg = ev.get("message", {})
            seen[msg.get("id", id(ev))] = msg
        elif t == "result":
            session_id = session_id or ev.get("session_id")
            is_error = bool(ev.get("is_error"))
    for msg in seen.values():
        for b in msg.get("content", []):
            if not isinstance(b, dict):
                continue
            if b.get("type") == "tool_use":
                tools.append(b.get("name", "").split("__")[-1])
            elif b.get("type") == "text" and b.get("text", "").strip():
                texts.append(b["text"].strip())
    return session_id, tools, is_error, " ".join(texts)


def _free_port(port: int) -> None:
    """Best-effort kill of any process still holding `port` from a previous
    (e.g. interrupted) run, so every run owns fresh servers."""
    try:
        out = subprocess.run(["lsof", "-ti", f"tcp:{port}"],
                             capture_output=True, text=True).stdout.split()
        for pid in out:
            try:
                os.kill(int(pid), signal.SIGKILL)
            except (ProcessLookupError, ValueError):
                pass
        if out:
            time.sleep(0.5)
    except FileNotFoundError:
        pass  # no lsof (non-mac/linux) — the per-run reset still guarantees clean state


def start_store() -> subprocess.Popen:
    _free_port(STORE_PORT)
    proc = subprocess.Popen(
        [sys.executable, "-m", "uvicorn", "sb.live.store_app:app", "--port", str(STORE_PORT),
         "--log-level", "warning"],
        cwd=str(REPO), stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    for _ in range(60):
        try:
            if httpx.get(f"{STORE_URL}/health", timeout=2).status_code == 200:
                return proc
        except Exception:
            pass
        time.sleep(0.5)
    proc.terminate()
    raise RuntimeError("store did not come up")


def start_mcp() -> subprocess.Popen:
    """Start the ONE persistent MCP server every claude turn attaches to."""
    _free_port(MCP_PORT)
    env = {**os.environ, "MCP_TRANSPORT": "streamable-http", "MCP_PORT": str(MCP_PORT),
           "STORE_BASE_URL": STORE_URL, "PYTHONPATH": str(REPO)}
    proc = subprocess.Popen(
        [sys.executable, "-m", "sb.live.mcp_app"],
        cwd=str(REPO), env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    for _ in range(60):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(1)
            if s.connect_ex(("127.0.0.1", MCP_PORT)) == 0:  # port is accepting connections
                return proc
        time.sleep(0.5)
    proc.terminate()
    raise RuntimeError("mcp server did not come up")


def run(model: str, seed: int, start: date, n_days: int, limit: Optional[int],
        per_turn_timeout: float = 240.0) -> None:
    corpus = load_corpus("corpus")
    plan = build_plan(corpus, start_date=start, seed=seed, n_days=n_days)
    order = [e for day in plan.per_day for e in day]
    if limit:
        order = order[:limit]

    store = start_store()
    httpx.post(f"{STORE_URL}/reset", timeout=10)
    mcp = start_mcp()
    mcp_config = _mcp_config()
    session: Optional[str] = None
    results: dict[str, object] = {}

    print(f"\n{BOLD}{BLUE}╔═══ SecretaryBench · live run ═══╗{RESET}")
    print(f"{BLUE}║{RESET} model {BOLD}{model}{RESET}")
    print(f"{BLUE}║{RESET} seed {seed} · {len(order)} emails · start {start}")
    print(f"{BOLD}{BLUE}╚═════════════════════════════════╝{RESET}")
    try:
        for i, eid in enumerate(order, 1):
            email = corpus.emails[eid]
            sd = plan.serve_date[eid]
            ctx = Context(sd, plan.anchors)
            rendered = resolver.render_body(email.body, ctx).text

            httpx.post(f"{STORE_URL}/inbox", timeout=10, json={
                "email_id": eid, "sender": email.sender, "recipients": email.recipients,
                "subject": email.subject, "body": rendered, "served_date": sd.isoformat()})

            user_msg = (
                f"Email-Id: {eid}\n"
                f"Today: {sd.strftime('%A, %B %d, %Y')}\n"
                f"From: {email.sender}\nTo: {', '.join(email.recipients)}\n"
                f"Subject: {email.subject}\n\n{rendered}")

            before = _all_ids(httpx.get(f"{STORE_URL}/state", timeout=10).json())

            cmd = ["claude", "-p", "--output-format", "stream-json", "--verbose",
                   "--tools", "", "--permission-mode", "bypassPermissions",
                   "--mcp-config", mcp_config, "--strict-mcp-config"]
            if session is None:
                cmd += ["--model", model, "--append-system-prompt", SYSTEM_PROMPT]
            else:
                cmd += ["--resume", session]
            cmd.append(user_msg)

            try:
                proc = subprocess.run(cmd, capture_output=True, text=True, timeout=per_turn_timeout)
                sid, tools, is_err, said = _parse_stream(proc.stdout)
                session = session or sid
                if is_err or proc.returncode != 0:
                    raise RuntimeError(proc.stderr.strip()[:200] or "claude error")
            except Exception as exc:
                results[eid] = None
                print(f"[{i:>2}] {eid:24s}  ERROR: {exc}")
                continue

            state = httpx.get(f"{STORE_URL}/state", timeout=10).json()
            new_ids = _all_ids(state) - before
            res = grade_email(email.answer, ctx,
                              _node_state(corpus, state, email.node, new_ids),
                              _turn_delta(corpus, state, new_ids))
            results[eid] = res
            _print_email(i, eid, sd.strftime("%a %b %d"), res, tools, said)
    finally:
        mcp.terminate()
        store.terminate()

    graded = [r for r in results.values() if r is not None]
    passed = sum(1 for r in graded if r.passed)
    errored = len(order) - len(graded)
    bar = "".join((f"{GREEN}●{RESET}" if (r and r.passed) else f"{RED}●{RESET}")
                  for r in results.values())
    pct = passed / len(order) if order else 0
    pct_col = GREEN if pct >= 0.8 else (YELLOW if pct >= 0.5 else RED)
    print(f"\n{BOLD}{BLUE}══════════════════════════════════{RESET}")
    print(f"  {bar}")
    print(f"  {BOLD}SCORE {pct_col}{passed}/{len(order)}{RESET} {BOLD}({pct:.0%}){RESET}"
          + (f"  {RED}{errored} errored{RESET}" if errored else ""))
    print(f"{BOLD}{BLUE}══════════════════════════════════{RESET}\n")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--model", default="claude-haiku-4-5")
    ap.add_argument("--seed", type=int, default=42)
    ap.add_argument("--start", default="2026-06-01")
    ap.add_argument("--days", type=int, default=60)
    ap.add_argument("--limit", type=int, default=None)
    a = ap.parse_args()
    run(a.model, a.seed, date.fromisoformat(a.start), a.days, a.limit)


if __name__ == "__main__":
    main()
