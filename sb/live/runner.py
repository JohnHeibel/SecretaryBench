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
import re
import shutil
import signal
import socket
import subprocess
import sys
import tempfile
import time
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Optional

import httpx

from sb import resolver
from sb.grader import NodeState, Obj, TurnDelta, grade_email
from sb.resolver import Context
from sb.schema import Corpus, load_corpus
from sb.scheduler import Levers, build_plan

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
    searched = f"  {CYAN}🔍 used search_inbox{RESET}" if "search_inbox" in tools else ""
    print(f"   {DIM}tools{RESET}  {GREY}{tool_str}{RESET}{searched}")
    if said:
        snippet = said[:200].replace("\n", " ")
        print(f"   {DIM}model said{RESET}  {YELLOW}“{snippet}…”{RESET}")


def _print_warnings(warnings: list[dict]) -> None:
    for w in warnings:
        details = ", ".join(f"{k}={v}" for k, v in w.items() if k not in {"kind", "today"})
        suffix = f" ({details})" if details else ""
        print(f"   {YELLOW}warning{RESET}  {w.get('kind', 'unknown')}{suffix}")


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


def _parse_stream(stdout: str) -> tuple[Optional[str], list[str], bool, str, Optional[str]]:
    """Return (session_id, tool_names_used, is_error, assistant_text, resolved_model).

    resolved_model is what the CLI actually loaded, not what we asked for. The init
    event has carried it all along; not reading it is how a silently-substituted
    model could look like a real score."""
    session_id, tools, is_error, texts = None, [], False, []
    resolved_model = None
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
            resolved_model = ev.get("model") or resolved_model
        elif t == "assistant":
            msg = ev.get("message", {})
            resolved_model = msg.get("model") or resolved_model
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
    return session_id, tools, is_error, " ".join(texts), resolved_model


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
    # Keep stderr: a crash here (bad mcp version, port in use) used to surface only as
    # the generic "did not come up", which hides the actual traceback.
    proc = subprocess.Popen(
        [sys.executable, "-m", "sb.live.mcp_app"],
        cwd=str(REPO), env=env, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True,
    )
    for _ in range(60):
        if proc.poll() is not None:      # died on startup — report why
            err = (proc.stderr.read() if proc.stderr else "").strip()
            raise RuntimeError(f"mcp server exited immediately:\n{err[-800:]}")
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(1)
            if s.connect_ex(("127.0.0.1", MCP_PORT)) == 0:  # port is accepting connections
                return proc
        time.sleep(0.5)
    proc.terminate()
    raise RuntimeError("mcp server did not come up within 30s")


# --- model drivers ---------------------------------------------------------
# One turn = one subprocess. `claude -p` and `codex exec` both attach to the SAME
# HTTP MCP server, keep ONE continuous session across days (--resume / exec resume),
# and stream JSON we parse for (session_id, tools_used, is_error, assistant_text).
# The store / scheduler / grader are identical for both -> apples-to-apples.

def _limit_reset_wait(text: str) -> Optional[int]:
    """Seconds to pause for a subscription usage-limit window, or None if this
    isn't one. Parses the CLI's "resets 2:20pm (America/Los_Angeles)" phrasing;
    falls back to 30min when the reset time is missing or unparseable."""
    if not re.search(r"(session|usage) limit", text, re.IGNORECASE):
        return None
    m = re.search(r"resets\s+(\d{1,2}):(\d{2})\s*(am|pm)(?:\s*\(([\w/_+-]+)\))?",
                  text, re.IGNORECASE)
    if not m:
        return 1800
    hh, mm, ap, tzname = int(m.group(1)), int(m.group(2)), m.group(3).lower(), m.group(4)
    hh = hh % 12 + (12 if ap == "pm" else 0)
    try:
        from zoneinfo import ZoneInfo
        tz = ZoneInfo(tzname) if tzname else None
    except Exception:
        tz = None
    now = datetime.now(tz)
    target = now.replace(hour=hh, minute=mm, second=0, microsecond=0)
    if target <= now:
        target += timedelta(days=1)
    return min(int((target - now).total_seconds()) + 120, 6 * 3600)


def infer_driver(model: str) -> str:
    """Pick the CLI that can serve a model id. OpenAI (gpt*/o1/o3/o4/codex) -> codex."""
    m = model.lower()
    if m.startswith(("gpt", "o1", "o3", "o4", "codex")) or "gpt" in m:
        return "codex"
    return "claude"


def _codex_flags(mcp_url: str, cwd: Optional[str], reasoning: Optional[str]) -> list[str]:
    flags = ["--json", "--dangerously-bypass-approvals-and-sandbox",
             "--skip-git-repo-check"]
    if cwd is not None:  # `codex exec resume` has no -C; subprocess cwd= isolates it anyway
        flags += ["-C", cwd]
    flags += ["-c", f'mcp_servers.secretary.url="{mcp_url}"',
              # the model must solve from its inbox only — no web/image tools
              "-c", "tools.web_search=false", "-c", "tools.image_gen=false"]
    if reasoning:
        flags += ["-c", f"model_reasoning_effort={reasoning}"]
    return flags


def build_cmd(driver: str, model: str, session: Optional[str], user_msg: str,
              mcp_config: str, mcp_url: str, codex_cwd: str,
              reasoning: Optional[str]) -> list[str]:
    if driver == "codex":
        if session is None:  # first turn carries the system prompt (codex has no flag for it)
            common = _codex_flags(mcp_url, codex_cwd, reasoning)
            return ["codex", "exec", *common, "-m", model, SYSTEM_PROMPT + "\n\n" + user_msg]
        common = _codex_flags(mcp_url, None, reasoning)
        return ["codex", "exec", "resume", *common, "-m", model, session, user_msg]
    # claude — NO `--tools ""` (in current CLI that means "zero tools available", which
    # blocked the MCP tools). bypassPermissions + strict-mcp-config already scope tools
    # to only our secretary MCP server; the subprocess runs in an isolated empty cwd.
    cmd = ["claude", "-p", "--output-format", "stream-json", "--verbose",
           "--permission-mode", "bypassPermissions",
           "--mcp-config", mcp_config, "--strict-mcp-config"]
    if session is None:
        cmd += ["--model", model, "--append-system-prompt", SYSTEM_PROMPT]
    else:
        cmd += ["--resume", session, "--model", model]
    cmd.append(user_msg)
    return cmd


def _parse_codex(stdout: str) -> tuple[Optional[str], list[str], bool, str, Optional[str]]:
    """Parse `codex exec --json` JSONL -> (thread_id, tool_names, is_error, text, model).
    Only thread_id (for resume) and is_error are load-bearing; tools/text are display."""
    session_id, tools, is_error, texts = None, [], False, []
    resolved_model = None
    for line in stdout.splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            ev = json.loads(line)
        except json.JSONDecodeError:
            continue
        t = ev.get("type")
        if t == "thread.started":
            session_id = ev.get("thread_id") or session_id
            resolved_model = ev.get("model") or resolved_model
        elif t in ("turn.failed", "error"):
            is_error = True
        elif t == "item.completed":
            item = ev.get("item", {}) if isinstance(ev.get("item"), dict) else {}
            itype = item.get("type") or item.get("item_type") or ""
            name = item.get("name") or item.get("tool") or ""
            if "tool" in itype or "call" in itype:
                if name:
                    tools.append(str(name).split("__")[-1])
            elif "message" in itype:
                txt = item.get("text") or item.get("content") or ""
                if isinstance(txt, str) and txt.strip():
                    texts.append(txt.strip())
    return session_id, tools, is_error, " ".join(texts), resolved_model


def parse_output(driver: str, stdout: str) -> tuple[Optional[str], list[str], bool, str, Optional[str]]:
    return _parse_codex(stdout) if driver == "codex" else _parse_stream(stdout)


def run(model: str, seed: int, start: date, n_days: int, limit: Optional[int],
        per_turn_timeout: float = 240.0, corpus_dir: str = "corpus",
        driver: str = "auto", levers: Optional[Levers] = None,
        reasoning: Optional[str] = None) -> None:
    if driver == "auto":
        driver = infer_driver(model)
    levers = levers or Levers()
    corpus = load_corpus(corpus_dir)
    plan = build_plan(corpus, start_date=start, seed=seed, n_days=n_days, levers=levers)
    # Isolated empty cwd for BOTH drivers: codex needs a trusted root; claude must not
    # ingest the repo's CLAUDE.md / answer-key files (context pollution + cheating).
    iso_cwd = tempfile.mkdtemp(prefix="sb_run_")

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
    print(f"{BLUE}║{RESET} model {BOLD}{model}{RESET} {DIM}(requested) via {driver}{RESET}")
    print(f"{BLUE}║{RESET} seed {seed} · {len(days)} days · {len(order)} emails · start {start}")
    print(f"{BLUE}║{RESET} levers daily {levers.daily_min}-{levers.daily_max} · urgency {levers.urgency_horizon}"
          f" · corpus {corpus_dir}")
    print(f"{BOLD}{BLUE}╚═════════════════════════════════╝{RESET}")
    email_no = 0
    resolved_model: Optional[str] = None
    try:
        for day_no, batch in enumerate(days, 1):
            sd = plan.serve_date[batch[0]]   # every email in a batch shares the serve day
            warning_cursor = len(httpx.get(f"{STORE_URL}/warnings", timeout=10).json())

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

            cmd = build_cmd(driver, model, session, user_msg, mcp_config, MCP_URL,
                            iso_cwd, reasoning)

            # Retry with exponential backoff: a transient rate-limit / overload
            # blip should pause-and-resume, not throw the day away. A subscription
            # session-limit window ("resets 2:20pm") can last hours — pause until
            # the stated reset instead of counting it against the retry budget.
            ok, sid, tools, said, detail = False, None, [], "", f"{driver} error"
            attempts = 7
            attempt = limit_pauses = 0
            while attempt < attempts:
                try:
                    # stdin=DEVNULL: codex exec blocks reading piped stdin otherwise;
                    # harmless for claude -p (it takes the prompt as an arg).
                    proc = subprocess.run(cmd, capture_output=True, text=True,
                                          timeout=per_turn_timeout, stdin=subprocess.DEVNULL,
                                          cwd=iso_cwd)
                    sid, tools, is_err, said, rmodel = parse_output(driver, proc.stdout)
                    if not is_err and proc.returncode == 0:
                        session = session or sid
                        # What the CLI actually loaded. Checked EVERY turn, not just the
                        # first: --resume could quietly drop back to a default model
                        # mid-run, and the score would still be filed under `model`.
                        if rmodel:
                            if resolved_model is None:
                                resolved_model = rmodel
                                match = rmodel == model or rmodel.startswith(model)
                                tag = f"{GREEN}✓{RESET}" if match else f"{RED}{BOLD}✗ MISMATCH{RESET}"
                                print(f"   {DIM}model served{RESET}  {BOLD}{rmodel}{RESET} {tag}")
                                if not match:
                                    print(f"   {RED}requested {model!r} but the CLI served "
                                          f"{rmodel!r} — this run does NOT measure {model}{RESET}")
                            elif rmodel != resolved_model:
                                print(f"   {RED}{BOLD}✗ MODEL DRIFT{RESET} {RED}day {day_no}: was "
                                      f"{resolved_model!r}, now {rmodel!r} — scores before and "
                                      f"after this point are not comparable{RESET}")
                                resolved_model = rmodel
                        ok = True
                        break
                    detail = (proc.stderr.strip() or proc.stdout.strip()[-160:])[:160] or f"{driver} returned is_error"
                except Exception as exc:
                    detail = str(exc)[:160]
                pause = _limit_reset_wait(f"{said} {detail}")
                if pause is not None and limit_pauses < 12:
                    limit_pauses += 1
                    print(f"day {day_no:>2} ({sd}): session limit — pausing "
                          f"{pause // 60}min until reset (pause {limit_pauses}/12)")
                    time.sleep(pause)
                    continue                      # a limit window is not a failed attempt
                attempt += 1
                if attempt < attempts:
                    wait = min(240, 15 * (2 ** (attempt - 1)))  # 15,30,60,120,240,240 (~12min total)
                    print(f"day {day_no:>2} ({sd}): transient ({detail}) — retry {attempt}/{attempts-1} in {wait}s")
                    time.sleep(wait)

            _print_day(day_no, sd.strftime("%a %b %d"), len(batch), tools, said)
            all_warnings = httpx.get(f"{STORE_URL}/warnings", timeout=10).json()
            _print_warnings(all_warnings[warning_cursor:])
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
        shutil.rmtree(iso_cwd, ignore_errors=True)

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
    # Stamp the config next to the score so a saved log can't be mislabelled later.
    print(f"  {DIM}served {resolved_model or model} · seed {seed} · daily "
          f"{levers.daily_min}-{levers.daily_max} · {len(order)} emails · {corpus_dir}{RESET}")
    print(f"{BOLD}{BLUE}══════════════════════════════════{RESET}\n")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--model", default="claude-haiku-4-5")
    ap.add_argument("--driver", default="auto", choices=["auto", "claude", "codex"],
                    help="which CLI drives the model (auto: infer from model id)")
    ap.add_argument("--seed", type=int, default=42)
    ap.add_argument("--start", default="2026-06-01")
    ap.add_argument("--days", type=int, default=60)
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--corpus", default="corpus")
    # scheduler levers — the recovered authored corpus needs a larger daily_max
    # (deadline-bearing emails cluster); default 5 matches the pilot.
    ap.add_argument("--daily-min", type=int, default=1)
    ap.add_argument("--daily-max", type=int, default=5)
    ap.add_argument("--urgency-horizon", type=int, default=7)
    ap.add_argument("--reasoning", default=None,
                    help="codex only: model_reasoning_effort (low|medium|high|xhigh)")
    ap.add_argument("--timeout", type=float, default=240.0, help="per-turn seconds")
    a = ap.parse_args()
    levers = Levers(daily_min=a.daily_min, daily_max=a.daily_max,
                    urgency_horizon=a.urgency_horizon)
    run(a.model, a.seed, date.fromisoformat(a.start), a.days, a.limit,
        per_turn_timeout=a.timeout, corpus_dir=a.corpus, driver=a.driver,
        levers=levers, reasoning=a.reasoning)


if __name__ == "__main__":
    main()
