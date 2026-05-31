"""
sb.resolver — the token & answer-key grammar evaluator.

This is the keystone of the benchmark (see ANSWER_KEY_GRAMMAR.md). It parses the
one governed date grammar and evaluates an expression against a serve anchor plus
a table of named anchors emitted by earlier emails. The SAME grammar drives both
the rendered email body and the grader's expected answer, so they cannot drift.

Public surface:
    resolve(expr, ctx)            -> date | datetime | Interval
    render_body(body, ctx)        -> RenderResult(text, emissions)
    human(value)                  -> str            (natural-language rendering)
    Context(serve, anchors)       -> evaluation context

Only DATES vary with serve order, so the grammar encodes dates only. Durations,
counts, names, times-of-day-as-prose are static and live in the answer key, not
here. The one exception carried in-grammar is an optional @HH:MM time attached to
a date (so a calendar event's date+time render from a single source).
"""
from __future__ import annotations

import calendar as _calmod
import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from typing import Union

Value = Union[date, datetime, "Interval"]

_WEEKDAYS = {"MON": 0, "TUE": 1, "WED": 2, "THU": 3, "FRI": 4, "SAT": 5, "SUN": 6}


class ResolverError(ValueError):
    """Raised on an unparseable or unresolvable expression."""


@dataclass(frozen=True)
class Interval:
    """An inclusive span of calendar days [start, end]."""
    start: date
    end: date

    def contains(self, d: date) -> bool:
        if isinstance(d, datetime):
            d = d.date()
        return self.start <= d <= self.end


@dataclass
class Context:
    serve: date                                   # the day the email is served (its "now")
    anchors: dict[str, Value] = field(default_factory=dict)
    facts: dict[str, str] = field(default_factory=dict)   # named static scalars (e.g. a duration)


@dataclass
class RenderResult:
    text: str
    emissions: dict[str, Value]
    fact_emissions: dict[str, str] = field(default_factory=dict)


# --- date helpers ----------------------------------------------------------

def _as_date(v: Value) -> date:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, Interval):
        return v.start
    return v


def add_months(d: date, n: int) -> date:
    """Shift a date by n calendar months, clamping the day to the month length."""
    month_index = d.month - 1 + n
    year = d.year + month_index // 12
    month = month_index % 12 + 1
    day = min(d.day, _calmod.monthrange(year, month)[1])
    return date(year, month, day)


def add_business_days(d: date, n: int) -> date:
    step = 1 if n >= 0 else -1
    remaining = abs(n)
    cur = d
    while remaining > 0:
        cur = cur + timedelta(days=step)
        if cur.weekday() < 5:        # Mon-Fri
            remaining -= 1
    return cur


def _week_of(d: date) -> Interval:
    monday = d - timedelta(days=d.weekday())
    return Interval(monday, monday + timedelta(days=6))


def _nth_weekday(year: int, month: int, weekday: int, n: int | str) -> date:
    first = date(year, month, 1)
    offset = (weekday - first.weekday()) % 7
    days = [first + timedelta(days=offset + 7 * k)
            for k in range(6)
            if (first + timedelta(days=offset + 7 * k)).month == month]
    if n == "last":
        return days[-1]
    if not (1 <= int(n) <= len(days)):
        raise ResolverError(f"no {n}th {weekday} in {year}-{month:02d}")
    return days[int(n) - 1]


# --- AST -------------------------------------------------------------------
# Each node is a small callable evaluated against a Context.

@dataclass
class _Serve:
    def eval(self, ctx: Context) -> Value:
        return ctx.serve


@dataclass
class _Anchor:
    name: str

    def eval(self, ctx: Context) -> Value:
        if self.name not in ctx.anchors:
            raise ResolverError(f"unknown anchor @{self.name}")
        return ctx.anchors[self.name]


@dataclass
class _NextWD:
    wd: int

    def eval(self, ctx: Context) -> Value:
        d = ctx.serve
        delta = (self.wd - d.weekday()) % 7
        delta = delta or 7                       # strictly after serve
        return d + timedelta(days=delta)


@dataclass
class _ThisWD:
    wd: int

    def eval(self, ctx: Context) -> Value:
        monday = ctx.serve - timedelta(days=ctx.serve.weekday())
        return monday + timedelta(days=self.wd)


@dataclass
class _NthWD:
    n: int | str
    wd: int
    month_off: int

    def eval(self, ctx: Context) -> Value:
        base = add_months(ctx.serve, self.month_off)
        return _nth_weekday(base.year, base.month, self.wd, self.n)


@dataclass
class _Dom:
    day: int
    month_off: int

    def eval(self, ctx: Context) -> Value:
        base = add_months(ctx.serve, self.month_off)
        last = _calmod.monthrange(base.year, base.month)[1]
        if not (1 <= self.day <= last):
            raise ResolverError(f"no day {self.day} in {base.year}-{base.month:02d}")
        return date(base.year, base.month, self.day)


@dataclass
class _MonthInterval:
    month_off: int

    def eval(self, ctx: Context) -> Value:
        base = add_months(ctx.serve, self.month_off)
        last = _calmod.monthrange(base.year, base.month)[1]
        return Interval(date(base.year, base.month, 1), date(base.year, base.month, last))


@dataclass
class _WeekOf:
    inner: object

    def eval(self, ctx: Context) -> Value:
        return _week_of(_as_date(self.inner.eval(ctx)))


@dataclass
class _Offset:
    base: object
    amount: int
    unit: str

    def eval(self, ctx: Context) -> Value:
        v = self.base.eval(ctx)
        d = _as_date(v)
        if self.unit == "d":
            nd = d + timedelta(days=self.amount)
        elif self.unit == "w":
            nd = d + timedelta(weeks=self.amount)
        elif self.unit == "bd":
            nd = add_business_days(d, self.amount)
        elif self.unit == "m":
            nd = add_months(d, self.amount)
        elif self.unit == "y":
            nd = add_months(d, self.amount * 12)
        else:
            raise ResolverError(f"bad offset unit {self.unit!r}")
        if isinstance(v, datetime):
            return datetime.combine(nd, v.time())
        return nd


@dataclass
class _AtTime:
    base: object
    hh: int
    mm: int

    def eval(self, ctx: Context) -> Value:
        d = _as_date(self.base.eval(ctx))
        return datetime(d.year, d.month, d.day, self.hh, self.mm)


# --- parser ----------------------------------------------------------------

_OFFSET_RE = re.compile(r"\s*([+-]\d+)(bd|d|w|m|y)")
_ATTACH_RE = re.compile(r"\s*@(\d{1,2}):(\d{2})\s*$")


def _parse_base(s: str) -> tuple[object, str]:
    """Parse a base term off the front of s; return (node, rest)."""
    s = s.lstrip()

    if s.startswith("serve"):
        return _Serve(), s[len("serve"):]

    m = re.match(r"@([A-Za-z_][A-Za-z0-9_]*)", s)
    if m:
        return _Anchor(m.group(1)), s[m.end():]

    m = re.match(r"next:([A-Za-z]{3})", s, re.I)
    if m:
        return _NextWD(_wd(m.group(1))), s[m.end():]

    m = re.match(r"this:([A-Za-z]{3})", s, re.I)
    if m:
        return _ThisWD(_wd(m.group(1))), s[m.end():]

    m = re.match(r"nth:(\d+|last),([A-Za-z]{3}),([+-]?\d+)m", s, re.I)
    if m:
        n = m.group(1).lower()
        n = "last" if n == "last" else int(n)
        return _NthWD(n, _wd(m.group(2)), int(m.group(3))), s[m.end():]

    m = re.match(r"dom:(\d+),([+-]?\d+)m", s, re.I)
    if m:
        return _Dom(int(m.group(1)), int(m.group(2))), s[m.end():]

    m = re.match(r"month:([+-]?\d+)m", s, re.I)
    if m:
        return _MonthInterval(int(m.group(1))), s[m.end():]

    if s.lower().startswith("week_of:"):
        rest = s[len("week_of:"):].lstrip()
        if not rest.startswith("("):
            raise ResolverError("week_of: requires a parenthesized sub-expression")
        depth, i = 0, 0
        for i, ch in enumerate(rest):
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth == 0:
                    break
        else:
            raise ResolverError("unbalanced parentheses in week_of:")
        inner = _parse_expr(rest[1:i])
        return _WeekOf(inner), rest[i + 1:]

    raise ResolverError(f"cannot parse base of {s!r}")


def _wd(token: str) -> int:
    try:
        return _WEEKDAYS[token.upper()]
    except KeyError:
        raise ResolverError(f"bad weekday {token!r}")


def _parse_expr(s: str) -> object:
    node, rest = _parse_base(s)
    # offsets
    while True:
        m = _OFFSET_RE.match(rest)
        if not m:
            break
        node = _Offset(node, int(m.group(1)), m.group(2))
        rest = rest[m.end():]
    # optional trailing time attach
    m = _ATTACH_RE.match(rest) or _ATTACH_RE.search(rest)
    if m and rest[:m.start()].strip() == "":
        node = _AtTime(node, int(m.group(1)), int(m.group(2)))
        rest = rest[m.end():]
    if rest.strip():
        raise ResolverError(f"trailing junk {rest!r} in expression")
    return node


# --- public API ------------------------------------------------------------

def resolve(expr: str, ctx: Context) -> Value:
    """Evaluate a grammar expression against ctx. Raises ResolverError."""
    return _parse_expr(expr.strip()).eval(ctx)


def value_kind(v: Value) -> str:
    if isinstance(v, datetime):
        return "datetime"
    if isinstance(v, Interval):
        return "interval"
    return "date"


_TOKEN_RE = re.compile(r"\{([^{}]*)\}")
_EMIT_RE = re.compile(r"^\s*!\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.+)$")
# a fact token: {=name = value} defines + renders a static scalar; {=name} references one
_FACT_RE = re.compile(r"^\s*=\s*([A-Za-z_][A-Za-z0-9_]*)\s*(?:=\s*(.+?))?\s*$")


def render_body(body: str, ctx: Context) -> RenderResult:
    """Render every {token} in an email body and collect emissions.

    Three token forms:
      {date_expr}            -> rendered to a concrete date string
      {!name = date_expr}    -> defines a named DATE anchor (and renders it)
      {=name = value}        -> defines a named static FACT scalar (and renders it)
      {=name}                -> references a previously-defined fact (and renders it)

    Emissions are evaluated against ctx and added to a local copy of the tables as
    we go, so a later token in the same body may reference an earlier emission.
    """
    emissions: dict[str, Value] = {}
    fact_emissions: dict[str, str] = {}
    local = Context(ctx.serve, dict(ctx.anchors), dict(ctx.facts))

    def _sub(match: re.Match) -> str:
        inner = match.group(1).strip()
        fact = _FACT_RE.match(inner)
        if fact:
            name, value = fact.group(1), fact.group(2)
            if value is not None:
                local.facts[name] = value.strip()
                fact_emissions[name] = value.strip()
            elif name not in local.facts:
                raise ResolverError(f"unknown fact ={name}")
            return human_fact(local.facts[name])
        emit = _EMIT_RE.match(inner)
        if emit:
            name, expr = emit.group(1), emit.group(2)
            value = resolve(expr, local)
            emissions[name] = value
            local.anchors[name] = value
            return human(value)
        return human(resolve(inner, local))

    text = _TOKEN_RE.sub(_sub, body)
    return RenderResult(text=text, emissions=emissions, fact_emissions=fact_emissions)


def resolve_scalar(ref: str, facts: dict[str, str]) -> str:
    """Resolve an answer scalar (duration/count). '@name' looks up a fact; anything
    else is a literal returned unchanged."""
    ref = str(ref).strip()
    if ref.startswith("@"):
        name = ref[1:]
        if name not in facts:
            raise ResolverError(f"unknown fact @{name}")
        return facts[name]
    return ref


def human_fact(value: str) -> str:
    """Natural-language rendering of a static fact (currently durations)."""
    v = value.strip()
    m = re.fullmatch(r"(\d+)\s*m", v)
    if m:
        return f"{m.group(1)}-minute"
    m = re.fullmatch(r"(\d+)\s*h", v)
    if m:
        return f"{m.group(1)}-hour"
    return v


def _ordinal(n: int) -> str:
    if 10 <= n % 100 <= 20:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


def human(value: Value) -> str:
    """Natural-language rendering used in email bodies."""
    if isinstance(value, Interval):
        return f"the week of {value.start.strftime('%B')} {_ordinal(value.start.day)}, {value.start.year}"
    if isinstance(value, datetime):
        h = value.hour % 12 or 12
        ampm = "AM" if value.hour < 12 else "PM"
        mm = "" if value.minute == 0 else f":{value.minute:02d}"
        day = f"{value.strftime('%A, %B')} {_ordinal(value.day)}, {value.year}"
        return f"{day} at {h}{mm} {ampm}"
    return f"{value.strftime('%A, %B')} {_ordinal(value.day)}, {value.year}"
