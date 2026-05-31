"""
authoring.chips — the chip <-> grammar compiler.

A "chip" is the non-technical, structured representation of a single date the
UI manipulates (a draggable pill). It compiles to exactly one expression in the
governed grammar (sb.resolver), and a grammar expression decompiles back into a
chip — so existing corpus files load straight into the editor and round-trip.

A chip is a plain dict:

    {
      "base":   { "kind": <BASE_KIND>, ...fields },   # required
      "offset": { "amount": int, "unit": "d|bd|w|m|y" } | null,
      "time":   { "hour": 0-23, "minute": 0-59 } | null,
      "emit_as": "anchor_name" | null                  # body chips only
    }

Base kinds (the curated palette) and the grammar base they compile to:

    serve                              -> serve
    anchor       {name}                -> @name
    next_weekday {weekday}             -> next:WD
    this_weekday {weekday}             -> this:WD
    nth_weekday  {n, weekday, month_offset}  -> nth:N,WD,+Mm   (n = 1..5 or "last")
    day_of_month {day, month_offset}   -> dom:D,+Mm
    month        {month_offset}        -> month:+Mm           (a whole-month interval)
    week_of      {inner: <chip>}       -> week_of:(...)
    raw          {token}               -> token verbatim       (escape hatch)

`compile_chip` never loses information; `parse_token` is best-effort and falls
back to a `raw` chip for anything it cannot cleanly decompose, so the escape
hatch always works.
"""
from __future__ import annotations

from typing import Any, Optional

from sb import resolver

# weekday <-> 3-letter token, in calendar order (Mon..Sun)
_WD_ORDER = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]
_WD_TO_INT = {w: i for i, w in enumerate(_WD_ORDER)}
_INT_TO_WD = {i: w for i, w in enumerate(_WD_ORDER)}

# friendly unit name <-> grammar suffix
UNIT_TO_SUFFIX = {"days": "d", "business-days": "bd", "weeks": "w", "months": "m", "years": "y"}
SUFFIX_TO_UNIT = {v: k for k, v in UNIT_TO_SUFFIX.items()}

BASE_KINDS = {
    "serve", "anchor", "next_weekday", "this_weekday",
    "nth_weekday", "day_of_month", "month", "week_of", "raw",
}


class ChipError(ValueError):
    """Raised when a chip spec is malformed and cannot be compiled."""


# --- compile: chip dict -> grammar string ----------------------------------

def _compile_base(base: dict) -> str:
    kind = base.get("kind")
    if kind not in BASE_KINDS:
        raise ChipError(f"unknown base kind {kind!r}")

    if kind == "serve":
        return "serve"

    if kind == "anchor":
        name = base.get("name")
        if not name:
            raise ChipError("anchor base needs a 'name'")
        return f"@{name}"

    if kind in ("next_weekday", "this_weekday"):
        wd = _wd_token(base.get("weekday"))
        return f"{'next' if kind == 'next_weekday' else 'this'}:{wd}"

    if kind == "nth_weekday":
        n = base.get("n")
        n_tok = "last" if str(n).lower() == "last" else int(n)
        wd = _wd_token(base.get("weekday"))
        mo = int(base.get("month_offset", 0))
        return f"nth:{n_tok},{wd},{_signed(mo)}m"

    if kind == "day_of_month":
        day = int(base["day"])
        mo = int(base.get("month_offset", 0))
        return f"dom:{day},{_signed(mo)}m"

    if kind == "month":
        mo = int(base.get("month_offset", 0))
        return f"month:{_signed(mo)}m"

    if kind == "week_of":
        inner = base.get("inner")
        if not inner:
            raise ChipError("week_of base needs an 'inner' chip")
        return f"week_of:({compile_chip(inner, allow_emit=False)})"

    if kind == "raw":
        token = (base.get("token") or "").strip()
        if not token:
            raise ChipError("raw base needs a non-empty 'token'")
        return token

    raise ChipError(f"unhandled base kind {kind!r}")  # pragma: no cover


def compile_chip(chip: dict, *, allow_emit: bool = True) -> str:
    """Compile a chip dict to its grammar expression (no surrounding braces)."""
    if not isinstance(chip, dict) or "base" not in chip:
        raise ChipError("chip must be a dict with a 'base'")

    # raw escape hatch passes through whole, ignoring offset/time wrappers
    if chip["base"].get("kind") == "raw":
        expr = _compile_base(chip["base"])
    else:
        expr = _compile_base(chip["base"])
        off = chip.get("offset")
        if off:
            amount = int(off["amount"])
            unit = off["unit"]
            if unit not in SUFFIX_TO_UNIT:
                unit = UNIT_TO_SUFFIX.get(unit, unit)
            if unit not in SUFFIX_TO_UNIT:
                raise ChipError(f"bad offset unit {off['unit']!r}")
            expr += f"{_signed(amount)}{unit}"
        t = chip.get("time")
        if t:
            expr += f"@{int(t['hour']):02d}:{int(t['minute']):02d}"

    # validate by round-tripping through the real parser
    try:
        resolver._parse_expr(expr.strip())
    except resolver.ResolverError as exc:
        raise ChipError(f"chip compiled to invalid grammar {expr!r}: {exc}")
    return expr


def compile_body_token(chip: dict) -> str:
    """Compile a chip into a body token string, including {} and optional emission.

    A chip with emit_as renders as {!name = expr}; otherwise {expr}.
    """
    expr = compile_chip(chip)
    emit = chip.get("emit_as")
    if emit:
        return f"{{!{emit} = {expr}}}"
    return f"{{{expr}}}"


# --- parse: grammar string -> chip dict (best-effort) -----------------------

def parse_token(expr: str, emit_as: Optional[str] = None) -> dict:
    """Decompile a grammar expression into a chip. Falls back to a raw chip."""
    try:
        ast = resolver._parse_expr(expr.strip())
        chip = _from_ast(ast)
    except (resolver.ResolverError, ChipError, Exception):  # noqa: BLE001 - any failure -> raw
        chip = {"base": {"kind": "raw", "token": expr.strip()}, "offset": None, "time": None}
    if emit_as:
        chip["emit_as"] = emit_as
    return chip


def _from_ast(node: Any) -> dict:
    """Walk the resolver AST into a normalized chip (time/offset/base)."""
    chip: dict = {"base": None, "offset": None, "time": None}

    # peel an optional trailing time-attach
    if isinstance(node, resolver._AtTime):
        chip["time"] = {"hour": node.hh, "minute": node.mm}
        node = node.base

    # peel a single offset (the curated palette allows one; nested -> raw)
    if isinstance(node, resolver._Offset):
        if isinstance(node.base, resolver._Offset):
            raise ChipError("stacked offsets -> raw")
        chip["offset"] = {"amount": node.amount, "unit": SUFFIX_TO_UNIT[node.unit]}
        node = node.base

    chip["base"] = _base_from_ast(node)
    return chip


def _base_from_ast(node: Any) -> dict:
    if isinstance(node, resolver._Serve):
        return {"kind": "serve"}
    if isinstance(node, resolver._Anchor):
        return {"kind": "anchor", "name": node.name}
    if isinstance(node, resolver._NextWD):
        return {"kind": "next_weekday", "weekday": _INT_TO_WD[node.wd]}
    if isinstance(node, resolver._ThisWD):
        return {"kind": "this_weekday", "weekday": _INT_TO_WD[node.wd]}
    if isinstance(node, resolver._NthWD):
        n = "last" if node.n == "last" else int(node.n)
        return {"kind": "nth_weekday", "n": n, "weekday": _INT_TO_WD[node.wd],
                "month_offset": node.month_off}
    if isinstance(node, resolver._Dom):
        return {"kind": "day_of_month", "day": node.day, "month_offset": node.month_off}
    if isinstance(node, resolver._MonthInterval):
        return {"kind": "month", "month_offset": node.month_off}
    if isinstance(node, resolver._WeekOf):
        return {"kind": "week_of", "inner": _from_ast(node.inner)}
    raise ChipError(f"cannot decompile {type(node).__name__}")


# --- small helpers ----------------------------------------------------------

def _wd_token(weekday: Any) -> str:
    if weekday is None:
        raise ChipError("weekday is required")
    w = str(weekday).strip().upper()[:3]
    if w not in _WD_TO_INT:
        raise ChipError(f"bad weekday {weekday!r}")
    return w


def _signed(n: int) -> str:
    return f"+{n}" if n >= 0 else str(n)
