"""
sb.live.runner — drive a REAL model over the corpus and grade it.

Architecture for one run:
    [this process] starts uvicorn store_app (persistent state)
        for each DAY in the plan:
            tell the store the date, POST the day's rendered emails to the inbox
            call `claude -p` (ONE continuous session, --resume) with a bodyless
                "it's <date>, you have new mail" prompt -> the model pulls its own
                inbox (list_new_emails -> get_email) and acts via MCP
            snapshot /state -> grade each email in the day's batch (objects are
                attributed back to emails by the email_id the model stamped on them)
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


def _print_day(day_no: int, served: str, n: int, tools: list[str], said: str) -> None:
    """One header per day. tools/said belong to the whole day's turn, not any one
    email, so they're reported here — not on the per-email rows."""
    print(f"\n{BOLD}{BLUE}── day {day_no} · {served} · {n} new email(s) ──{RESET}")
    tool_str = ", ".join(tools) if tools else "(none)"
    searched = f"  {CYAN}🔍 searched inbox{RESET}" if any("inbox" in t or t == "get_email" for t in tools) else ""
    print(f"   {DIM}tools{RESET}  {GREY}{tool_str}{RESET}{searched}")
    if said:
        snippet = said[:200].replace("\n", " ")
        print(f"   {DIM}model said{RESET}  {YELLOW}“{snippet}…”{RESET}")


def _print_email(i: int, eid: str, served: str, res) -> None:
    ok = res.passed
    badge = f"{GREEN}{BOLD} ✓ PASS {RESET}" if ok else f"{RED}{BOLD} ✗ FAIL {RESET}"
    print(f"\n{badge} {BOLD}[{i}] {eid}{RESET}  {GREY}· served {served}{RESET}")
    for d in res.details:
        tick = f"{GREEN}✓{RESET}" if d["passed"] else f"{RED}✗{RESET}"
        rc = GREEN if d["passed"] else RED
        print(f"   {tick} {DIM}expected{RESET}  {CYAN}{d['expected']}{RESET}")
        print(f"     {DIM}actual  {RESET}  {d['actual']}")
        print(f"     {DIM}why     {RESET}  {rc}{d['reason']}{RESET}")

SYSTEM_PROMPT = """You are an executive assistant managing one person's calendar and to-dos from their email inbox. You work ONE simulated day at a time, and you keep ONE continuous memory across all of them.

Each turn is a new day. You are told only the date and that new mail has arrived — you are NOT handed the emails. You must work your own inbox:

1. Call list_new_emails to see everything that arrived today (id, sender, subject — no body).
2. For each one, call get_email(email_id) to read it, then decide the correct action(s) and call the matching tool — or do nothing.
3. Work through the whole day's batch, then stop. Do not narrate.

TOOLS:
- list_new_emails — today's arrivals (ids + subjects). Always start here.
- get_email(email_id) — read the full body of one email (today's or any earlier one).
- create_event / update_event / delete_event — the calendar.
- create_todo / update_todo / delete_todo — tasks that have a deadline.
- search_inbox — look back at earlier emails. USE THIS whenever an email refers to a fact, date, policy, person, or decision you don't fully remember (e.g. "two weeks after the signing", "per the new meeting policy"). The information was usually set in an earlier email.
- (no tool) — for FYIs, newsletters, confirmations, or anything needing no calendar/to-do change. Over-acting is the most common mistake; when in doubt, do nothing.

EVERY create_event and create_todo REQUIRES the email_id of the email you are acting on — the id you got from list_new_emails. Copy it exactly.

DATES: the turn gives you "Today: <date>". Compute all relative dates from it. Use ISO 8601 in tool calls. Defaults unless the email says otherwise: events 9am, 1 hour long; to-dos due 5pm.

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
        per_turn_timeout: float = 240.0, corpus_dir: str = "corpus") -> None:
    corpus = load_corpus(corpus_dir)
    plan = build_plan(corpus, start_date=start, seed=seed, n_days=n_days)

    # One turn = one day. Drop empty days (no mail arrived) and, if --limit is set,
    # truncate so the total number of emails graded stays under the cap.
    days: list[list[str]] = []
    seen = 0
    for batch in plan.per_day:
        if not batch:
            continue
        b = batch[: (limit - seen)] if limit else batch
        if b:
            days.append(b)
            seen += len(b)
        if limit and seen >= limit:
            break
    order = [e for b in days for e in b]

    store = start_store()
    httpx.post(f"{STORE_URL}/reset", timeout=10)
    mcp = start_mcp()
    mcp_config = _mcp_config()
    session: Optional[str] = None
    results: dict[str, object] = {}

    print(f"\n{BOLD}{BLUE}╔═══ SecretaryBench · live run ═══╗{RESET}")
    print(f"{BLUE}║{RESET} model {BOLD}{model}{RESET}")
    print(f"{BLUE}║{RESET} seed {seed} · {len(days)} days · {len(order)} emails · start {start}")
    print(f"{BOLD}{BLUE}╚═════════════════════════════════╝{RESET}")
    email_no = 0
    try:
        for day_no, batch in enumerate(days, 1):
            sd = plan.serve_date[batch[0]]   # every email in a batch shares the serve day

            # Tell the store which day it is, then drop today's mail into the inbox so
            # list_new_emails can see it. Bodies go to the store, NOT the prompt.
            httpx.post(f"{STORE_URL}/day", timeout=10, json={"date": sd.isoformat()})
            for eid in batch:
                email = corpus.emails[eid]
                rendered = resolver.render_body(email.body, Context(sd, plan.anchors)).text
                httpx.post(f"{STORE_URL}/inbox", timeout=10, json={
                    "email_id": eid, "sender": email.sender, "recipients": email.recipients,
                    "subject": email.subject, "body": rendered, "served_date": sd.isoformat()})

            user_msg = (
                f"Today: {sd.strftime('%A, %B %d, %Y')}\n"
                f"New mail has arrived. List your inbox, work through everything that "
                f"arrived today, then stop.")

            before = _all_ids(httpx.get(f"{STORE_URL}/state", timeout=10).json())

            cmd = ["claude", "-p", "--output-format", "stream-json", "--verbose",
                   "--tools", "", "--permission-mode", "bypassPermissions",
                   "--mcp-config", mcp_config, "--strict-mcp-config"]
            if session is None:
                cmd += ["--model", model, "--append-system-prompt", SYSTEM_PROMPT]
            else:
                cmd += ["--resume", session]
            cmd.append(user_msg)

            # Retry with exponential backoff: a transient rate-limit / overload
            # blip should pause-and-resume, not throw the day away.
            ok, sid, tools, said, detail = False, None, [], "", "claude error"
            for attempt in range(5):
                try:
                    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=per_turn_timeout)
                    sid, tools, is_err, said = _parse_stream(proc.stdout)
                    if not is_err and proc.returncode == 0:
                        session = session or sid
                        ok = True
                        break
                    detail = proc.stderr.strip()[:160] or "claude returned is_error"
                except Exception as exc:
                    detail = str(exc)[:160]
                if attempt < 4:
                    wait = 5 * (3 ** attempt)        # 5s, 15s, 45s, 135s
                    print(f"day {day_no:>2} ({sd}): transient ({detail}) — retry {attempt+1}/4 in {wait}s")
                    time.sleep(wait)

            _print_day(day_no, sd.strftime("%a %b %d"), len(batch), tools, said)
            if not ok:
                for eid in batch:
                    email_no += 1
                    results[eid] = None
                    print(f"[{email_no:>2}] {eid:24s}  ERROR after retries: {detail}")
                continue

            # One turn handled the whole day; split its new objects back to each email
            # by the email_id the model stamped on them.
            state = httpx.get(f"{STORE_URL}/state", timeout=10).json()
            day_new = _all_ids(state) - before
            by_eid = {r["id"]: r.get("email_id", "") for r in state["events"] + state["todos"]}
            for eid in batch:
                email_no += 1
                email = corpus.emails[eid]
                ctx = Context(plan.serve_date[eid], plan.anchors)
                eid_new = {i for i in day_new if by_eid.get(i) == eid}
                res = grade_email(email.answer, ctx,
                                  _node_state(corpus, state, email.node, eid_new),
                                  _turn_delta(corpus, state, eid_new))
                results[eid] = res
                _print_email(email_no, eid, sd.strftime("%a %b %d"), res)
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
    ap.add_argument("--corpus", default="corpus")
    a = ap.parse_args()
    run(a.model, a.seed, date.fromisoformat(a.start), a.days, a.limit, corpus_dir=a.corpus)


if __name__ == "__main__":
    main()
