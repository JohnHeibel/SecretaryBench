"""Centralized, structured logging for SecretaryBench.

All simulation output flows through this module so the format stays
consistent across engine, model_runner, and grader.  Output is both
human-readable (box-drawing, ANSI color) and LLM-parseable (tagged
sections with key: value pairs).

Colour is auto-disabled when stdout is not a TTY (piped / redirected).
"""

from __future__ import annotations

import json
import os
import sys
import threading
from datetime import datetime, timezone
from typing import Any


# ── ANSI colour ──────────────────────────────────────────────────────────

_USE_COLOR = hasattr(sys.stdout, "isatty") and sys.stdout.isatty()

if os.environ.get("NO_COLOR"):
    _USE_COLOR = False
if os.environ.get("FORCE_COLOR"):
    _USE_COLOR = True

_RST  = "\033[0m"  if _USE_COLOR else ""
_BOLD = "\033[1m"  if _USE_COLOR else ""
_DIM  = "\033[2m"  if _USE_COLOR else ""
_CYN  = "\033[36m" if _USE_COLOR else ""
_GRN  = "\033[32m" if _USE_COLOR else ""
_YLW  = "\033[33m" if _USE_COLOR else ""
_RED  = "\033[31m" if _USE_COLOR else ""
_MAG  = "\033[35m" if _USE_COLOR else ""
_BLU  = "\033[34m" if _USE_COLOR else ""
_WHT  = "\033[97m" if _USE_COLOR else ""


# ── Box-drawing primitives ───────────────────────────────────────────────

_H  = "─"
_V  = "│"
_TL = "┌"
_TR = "┐"
_BL = "└"
_BR = "┘"
_LT = "├"
_RT = "┤"


def _box_top(width: int = 60) -> str:
    return f"{_DIM}{_TL}{_H * width}{_TR}{_RST}"


def _box_bot(width: int = 60) -> str:
    return f"{_DIM}{_BL}{_H * width}{_BR}{_RST}"


def _box_sep(width: int = 60) -> str:
    return f"{_DIM}{_LT}{_H * width}{_RT}{_RST}"


def _box_row(text: str, width: int = 60) -> str:
    plain_len = len(_strip_ansi(text))
    inner = width - 2
    pad = max(inner - plain_len, 0)
    return f"{_DIM}{_V}{_RST} {text}{' ' * pad} {_DIM}{_V}{_RST}"


def _strip_ansi(s: str) -> str:
    import re
    return re.sub(r"\033\[[0-9;]*m", "", s)


# ── Timestamp helper ─────────────────────────────────────────────────────

def _ts() -> str:
    return datetime.now().strftime("%H:%M:%S")


# ── Public API ───────────────────────────────────────────────────────────

W = 68


def sim_header(n_scenarios: int, sim_days: int, start_date: str, model: str = "") -> None:
    """Print the simulation banner at startup."""
    print()
    print(_box_top(W))
    print(_box_row(f"{_BOLD}{_CYN}SecretaryBench{_RST}  {_DIM}simulation run{_RST}", W))
    print(_box_sep(W))
    print(_box_row(f"{_DIM}scenarios :{_RST}  {_WHT}{n_scenarios}{_RST}", W))
    print(_box_row(f"{_DIM}sim days  :{_RST}  {_WHT}{sim_days}{_RST}", W))
    print(_box_row(f"{_DIM}start     :{_RST}  {_WHT}{start_date}{_RST}", W))
    if model:
        print(_box_row(f"{_DIM}model     :{_RST}  {_WHT}{model}{_RST}", W))
    print(_box_row(f"{_DIM}begin     :{_RST}  {_WHT}{_ts()}{_RST}", W))
    print(_box_bot(W))
    print()


def day_header(day_num: int, date_str: str, n_emails: int) -> None:
    """Print a day separator with email count."""
    label = f"{_BOLD}{_BLU}DAY {day_num}{_RST}  {_DIM}{date_str}{_RST}"
    emails = f"{_CYN}{n_emails}{_RST} email(s)"
    print(f"  {_DIM}{_H * 3}{_RST} {label}  {_DIM}{_H * 3}{_RST}  {emails}")


def email_served(scenario_type: str, scenario_id: str, email_idx: int,
                 total_emails: int, subject: str) -> None:
    """Log a single email being served to the model."""
    tag = f"{_MAG}{scenario_type}{_RST}"
    pos = f"{_DIM}{email_idx}/{total_emails}{_RST}"
    subj = f"{_WHT}{subject[:50]}{_RST}" if subject else ""
    print(f"      {_DIM}{_V}{_RST} {tag} {scenario_id}  {pos}  {subj}")


def model_call(scenario_id: str, round_num: int) -> None:
    """Log that the model is being called."""
    print(f"      {_DIM}{_V} model  round {round_num}  scenario {scenario_id}{_RST}")


def tool_use(tool_name: str, is_error: bool = False) -> None:
    """Log a tool call from the model."""
    status = f"{_RED}ERR{_RST}" if is_error else f"{_GRN}OK{_RST}"
    print(f"      {_DIM}{_V}   tool  {_RST}{_YLW}{tool_name}{_RST}  [{status}]")


def grade_result(scenario_type: str, scenario_id: str,
                 score: int, max_score: int,
                 details: list[dict] | None = None) -> None:
    """Log a grading result for a completed scenario."""
    if max_score == 0:
        return
    pct = (100.0 * score / max_score) if max_score else 0.0
    color = _GRN if pct == 100 else _YLW if pct >= 50 else _RED
    tag = f"{_MAG}{scenario_type}{_RST}"
    result = f"{color}{score}/{max_score}{_RST}  {_DIM}({pct:.0f}%){_RST}"
    print(f"      {_DIM}{_V}{_RST} {_CYN}grade{_RST}  {tag} {scenario_id}  {result}")
    if details:
        for d in details:
            mark = f"{_GRN}✓{_RST}" if d["passed"] else f"{_RED}✗{_RST}"
            crit = d.get("criteria", "")
            print(f"      {_DIM}{_V}         {mark} {_DIM}{crit}{_RST}")


def day_summary(day_num: int, score: int, max_score: int) -> None:
    """Print end-of-day score line."""
    if max_score == 0:
        return
    pct = 100.0 * score / max_score
    color = _GRN if pct == 100 else _YLW if pct >= 50 else _RED
    print(f"      {_DIM}{_BL}{_H * 2}{_RST} day {day_num} total: "
          f"{color}{score}/{max_score}{_RST}\n")


def sim_footer(total_score: int, total_max: int,
               inactive: int, active: int,
               email_score: int = 0, email_max: int = 0) -> None:
    """Print the simulation completion banner.

    The scenario-level score (total_score/total_max) is the headline. The
    email-level score, when provided, prints on a second line for contrast —
    a scenario can score 1/1 with email failures, or 0/1 with most emails
    correct, and the gap is diagnostically useful.
    """
    pct = (100.0 * total_score / total_max) if total_max else 0.0
    color = _GRN if pct >= 80 else _YLW if pct >= 50 else _RED

    print()
    print(_box_top(W))
    print(_box_row(f"{_BOLD}{_CYN}Simulation Complete{_RST}", W))
    print(_box_sep(W))
    print(_box_row(f"{_DIM}scenario  :{_RST}  {color}{_BOLD}{total_score}/{total_max}{_RST}  {_DIM}({pct:.1f}%){_RST}", W))
    if email_max:
        epct = 100.0 * email_score / email_max
        ecolor = _GRN if epct >= 80 else _YLW if epct >= 50 else _RED
        print(_box_row(f"{_DIM}per-email :{_RST}  {ecolor}{_BOLD}{email_score}/{email_max}{_RST}  {_DIM}({epct:.1f}%){_RST}", W))
    print(_box_row(f"{_DIM}remaining :{_RST}  {_WHT}{inactive}{_RST} inactive  {_WHT}{active}{_RST} active", W))
    print(_box_row(f"{_DIM}end       :{_RST}  {_WHT}{_ts()}{_RST}", W))
    print(_box_bot(W))
    print()


def type_breakdown(by_type: dict[str, dict[str, int]],
                   title: str = "Score Breakdown by Type") -> None:
    """Print a score-by-scenario-type table, sorted worst-first."""
    if not by_type:
        return

    rows: list[tuple[str, int, int, int, int, float]] = []
    for stype, b in by_type.items():
        lost = b["max_score"] - b["score"]
        pct = (100.0 * b["score"] / b["max_score"]) if b["max_score"] else 0.0
        rows.append((stype, b["count"], b["score"], b["max_score"], lost, pct))
    rows.sort(key=lambda r: (-r[4], -r[3]))

    hdr_type  = f"{_BOLD}{'type':<10}{_RST}"
    hdr_n     = f"{_BOLD}{'n':>4}{_RST}"
    hdr_score = f"{_BOLD}{'score':>7}{_RST}"
    hdr_max   = f"{_BOLD}{'max':>7}{_RST}"
    hdr_lost  = f"{_BOLD}{'lost':>6}{_RST}"
    hdr_pct   = f"{_BOLD}{'pct':>7}{_RST}"

    print(_box_top(W))
    print(_box_row(f"{_BOLD}{_CYN}{title}{_RST}", W))
    print(_box_sep(W))
    print(_box_row(f"{hdr_type} {hdr_n} {hdr_score} {hdr_max} {hdr_lost} {hdr_pct}", W))
    print(_box_sep(W))

    for stype, n, score, mx, lost, pct in rows:
        color = _GRN if pct == 100 else _YLW if pct >= 50 else _RED
        line = (f"{_MAG}{stype:<10}{_RST} {n:>4} "
                f"{color}{score:>7}{_RST} {mx:>7} "
                f"{_RED if lost else _DIM}{lost:>6}{_RST} "
                f"{color}{pct:>6.1f}%{_RST}")
        print(_box_row(line, W))

    print(_box_bot(W))
    print()


def model_summary(stats: dict[str, Any], model_name: str = "",
                  token_log: str = "", tool_log: str = "") -> None:
    """Print the model_runner exit summary."""
    total_calls = sum(stats.get("tool_calls", {}).values())
    scenarios   = stats.get("scenarios_run", 0)
    rounds      = stats.get("rounds_total", 0)
    in_tok      = stats.get("input_tokens", 0)
    out_tok     = stats.get("output_tokens", 0)
    total_tok   = in_tok + out_tok

    print()
    print(_box_top(W))
    print(_box_row(f"{_BOLD}{_CYN}Model Runner Summary{_RST}", W))
    print(_box_sep(W))

    def _kv(k: str, v: Any) -> None:
        print(_box_row(f"{_DIM}{k:<18}:{_RST}  {_WHT}{v}{_RST}", W))

    _kv("model", model_name)
    _kv("scenarios run", scenarios)
    _kv("calendars created", stats.get("calendars_created", 0))
    _kv("total tool calls", total_calls)
    _kv("tool errors", stats.get("tool_errors", 0))
    _kv("turn failures", stats.get("turn_failures", 0))
    _kv("api rounds", rounds)

    print(_box_sep(W))
    print(_box_row(f"{_BOLD}Tokens{_RST}", W))
    print(_box_sep(W))
    _kv("input", f"{in_tok:,}")
    _kv("output", f"{out_tok:,}")
    _kv("total", f"{total_tok:,}")
    if scenarios:
        _kv("avg tokens/turn", f"{total_tok / scenarios:,.0f}")
        _kv("avg rounds/turn", f"{rounds / scenarios:.2f}")

    tool_calls = stats.get("tool_calls", {})
    if tool_calls:
        print(_box_sep(W))
        print(_box_row(f"{_BOLD}Tool Call Breakdown{_RST}", W))
        print(_box_sep(W))
        for name, count in sorted(tool_calls.items(), key=lambda x: -x[1]):
            bar_len = min(count, 30)
            bar = f"{_CYN}{'█' * bar_len}{_RST}"
            print(_box_row(f"  {_YLW}{name:<22}{_RST} {count:>4}  {bar}", W))

    if token_log or tool_log:
        print(_box_sep(W))
        print(_box_row(f"{_BOLD}Log Files{_RST}", W))
        print(_box_sep(W))
        if token_log:
            _kv("token log", token_log)
        if tool_log:
            _kv("tool log", tool_log)

    print(_box_bot(W))
    print()


def error(component: str, msg: str, detail: str = "") -> None:
    """Log an error to stderr."""
    prefix = f"{_RED}{_BOLD}ERROR{_RST} {_DIM}[{component}]{_RST}"
    sys.stderr.write(f"{prefix} {msg}\n")
    if detail:
        for line in detail.strip().split("\n"):
            sys.stderr.write(f"  {_DIM}{line}{_RST}\n")


def warn(component: str, msg: str) -> None:
    """Log a warning."""
    prefix = f"{_YLW}WARN{_RST}  {_DIM}[{component}]{_RST}"
    print(f"{prefix} {msg}")


def info(component: str, msg: str) -> None:
    """Log an info message."""
    prefix = f"{_BLU}INFO{_RST}  {_DIM}[{component}]{_RST}"
    print(f"{prefix} {msg}")


# ── Delivery log (JSONL) ─────────────────────────────────────────────────

# Path for the email-delivery JSONL log. Naming matches Miguel's
# token_usage.jsonl / tool_calls.jsonl pattern. Empty string disables.
DELIVERY_LOG_PATH = os.environ.get("DELIVERY_LOG_PATH", "delivery_log.jsonl")
_delivery_log_lock = threading.Lock()


def log_delivery(entry: dict) -> None:
    """Append one row to the delivery JSONL log. Lock-guarded for parallel runs.

    Failures are swallowed — logging must never break a benchmark run.
    """
    if not DELIVERY_LOG_PATH:
        return
    row = {"ts": datetime.now(timezone.utc).isoformat(), **entry}
    try:
        with _delivery_log_lock, open(DELIVERY_LOG_PATH, "a") as f:
            f.write(json.dumps(row) + "\n")
    except OSError:
        pass


# ── Pool transition trace ────────────────────────────────────────────────
#
# Off by default — verbose runs already print plenty. Set BENCH_TRACE_POOL=1
# to get a play-by-play of inactive->active->ready->served transitions. ~280
# extra lines per full run; useful for diagnosing scheduling bugs.

_TRACE_POOL = os.environ.get("BENCH_TRACE_POOL", "") == "1"


def pool_activate(scenario_type: str, scenario_id: str, day: int) -> None:
    """One-line trace when a scenario moves inactive -> active."""
    if not _TRACE_POOL:
        return
    print(f"      {_DIM}{_V} pool   activate {_RST}{_MAG}{scenario_type}{_RST} "
          f"{scenario_id} {_DIM}(day {day}){_RST}")


def pool_email_ready(scenario_type: str, scenario_id: str,
                     email_idx: int, total: int, day: int) -> None:
    """One-line trace when an email transitions to READY_TO_SERVE."""
    if not _TRACE_POOL:
        return
    print(f"      {_DIM}{_V} pool   ready    {_RST}{_MAG}{scenario_type}{_RST} "
          f"{scenario_id} {_DIM}email {email_idx}/{total} (day {day}){_RST}")


def pool_email_served(scenario_type: str, scenario_id: str,
                      email_idx: int, total: int, day: int) -> None:
    """One-line trace when an email transitions to SERVED."""
    if not _TRACE_POOL:
        return
    print(f"      {_DIM}{_V} pool   served   {_RST}{_MAG}{scenario_type}{_RST} "
          f"{scenario_id} {_DIM}email {email_idx}/{total} (day {day}){_RST}")
