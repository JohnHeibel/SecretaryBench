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

Dates vary with serve order, so the grammar computes them. A date may carry an
optional time-of-day suffix — `@HH:MM` (a point) or `@HH:MM-HH:MM` (a within-day
span) — which is static (it does not shift with serve order) but rides INSIDE the
one governed expression, so the email body and the answer key still render from a
single source and cannot drift. Clock times are bounded by the CEO's work hours
(WORK_START..WORK_END); a token resolving outside them is a grammar error, since
such a slot can never be booked. A bare date still resolves to a whole DAY and is
graded day-level; a timed expression resolves to a datetime / TimeInterval and is
graded at minute granularity (start, and end/duration for an interval).
"""
from __future__ import annotations

import calendar as _calmod
import re
from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta
from typing import Union

Value = Union[date, datetime, "Interval", "TimeInterval"]

_WEEKDAYS = {"MON": 0, "TUE": 1, "WED": 2, "THU": 3, "FRI": 4, "SAT": 5, "SUN": 6}

# The CEO's bookable window. A clock time outside it is never schedulable, so a
# time token resolving outside [WORK_START, WORK_END] is rejected at parse time
# (and therefore at lint). Phase-2 free-slot search treats this window as the
# per-day universe of bookable time.
WORK_START = time(5, 0)
WORK_END = time(23, 0)


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


@dataclass(frozen=True)
class TimeInterval:
    """A within-day span [start, end) with clock times — what a timed obligation
    occupies. Distinct from Interval (whole days). start and end are datetimes on
    the same calendar day, with start < end."""
    start: datetime
    end: datetime

    def overlaps(self, other: "TimeInterval") -> bool:
        # half-open [start, end): touching edges (end == other.start) do NOT overlap.
        return self.start < other.end and other.start < self.end


@dataclass
class Context:
    serve: date                                   # the day the email is served (its "now")
    anchors: dict[str, Value] = field(default_factory=dict)


@dataclass
class RenderResult:
    text: str
    emissions: dict[str, Value]


# --- date helpers ----------------------------------------------------------

def _as_date(v: Value) -> date:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, TimeInterval):
        return v.start.date()
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
    ref: object = None                           # None -> serve-relative; else anchor expr

    def eval(self, ctx: Context) -> Value:
        d = ctx.serve if self.ref is None else _as_date(self.ref.eval(ctx))
        delta = (self.wd - d.weekday()) % 7
        delta = delta or 7                       # strictly after the reference date
        return d + timedelta(days=delta)


@dataclass
class _ThisWD:
    wd: int
    ref: object = None                           # None -> serve-relative; else anchor expr

    def eval(self, ctx: Context) -> Value:
        d = ctx.serve if self.ref is None else _as_date(self.ref.eval(ctx))
        monday = d - timedelta(days=d.weekday())
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


def _apply_offset(d: date, amount: int, unit: str) -> date:
    if unit == "d":
        return d + timedelta(days=amount)
    if unit == "w":
        return d + timedelta(weeks=amount)
    if unit == "bd":
        return add_business_days(d, amount)
    if unit == "m":
        return add_months(d, amount)
    if unit == "y":
        return add_months(d, amount * 12)
    raise ResolverError(f"bad offset unit {unit!r}")


@dataclass
class _Offset:
    base: object
    amount: int
    unit: str

    def eval(self, ctx: Context) -> Value:
        v = self.base.eval(ctx)
        if isinstance(v, TimeInterval):       # offset an anchored time-span by whole days, keep clock
            nd = _apply_offset(v.start.date(), self.amount, self.unit)
            return TimeInterval(datetime.combine(nd, v.start.time()),
                                datetime.combine(nd, v.end.time()))
        d = _as_date(v)
        nd = _apply_offset(d, self.amount, self.unit)
        if isinstance(v, datetime):
            return datetime.combine(nd, v.time())
        return nd


@dataclass
class _TimeAttach:
    """A date expression with a clock time attached: `<dateexpr> @HH:MM` -> datetime,
    `<dateexpr> @HH:MM-HH:MM` -> TimeInterval. The date part is serve-relative; the
    clock time is a static literal validated against work hours at parse time."""
    base: object
    start: tuple            # (hour, minute)
    end: tuple | None       # (hour, minute) for an interval, else None

    def eval(self, ctx: Context) -> Value:
        d = _as_date(self.base.eval(ctx))
        start_dt = datetime.combine(d, time(self.start[0], self.start[1]))
        if self.end is None:
            return start_dt
        return TimeInterval(start_dt, datetime.combine(d, time(self.end[0], self.end[1])))


# --- parser ----------------------------------------------------------------

_OFFSET_RE = re.compile(r"\s*([+-]\d+)(bd|d|w|m|y)")
_FROM_RE = re.compile(r"\s+from\s+", re.I)
_TIME_RE = re.compile(r"\s*@(\d{1,2}):(\d{2})(?:\s*-\s*(\d{1,2}):(\d{2}))?")


def _parse_time(hh: str, mm: str) -> tuple:
    h, m = int(hh), int(mm)
    if not (0 <= h <= 23 and 0 <= m <= 59):
        raise ResolverError(f"bad clock time {h:02d}:{m:02d} (expected 00:00–23:59)")
    return (h, m)


def _hhmm(t: tuple) -> str:
    return f"{t[0]:02d}:{t[1]:02d}"


def _check_work_hours(start: tuple, end: tuple | None) -> None:
    st = time(start[0], start[1])
    if st < WORK_START or st > WORK_END:
        raise ResolverError(
            f"time {_hhmm(start)} is outside work hours "
            f"{WORK_START.strftime('%H:%M')}–{WORK_END.strftime('%H:%M')}")
    if end is not None:
        et = time(end[0], end[1])
        if et <= st:
            raise ResolverError(f"end {_hhmm(end)} must be after start {_hhmm(start)}")
        if et > WORK_END:
            raise ResolverError(
                f"end {_hhmm(end)} is past work hours end {WORK_END.strftime('%H:%M')}")


def _parse_from_clause(rest: str) -> tuple[object, str]:
    """If rest begins with a ` from ` clause, parse the remainder of the string
    as a full grammar sub-expression (the anchor reference point) and return
    (ref_node, ""). Otherwise return (None, rest) leaving rest for offsets."""
    m = _FROM_RE.match(rest)
    if not m:
        return None, rest
    return _parse_expr(rest[m.end():]), ""


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
        ref, rest = _parse_from_clause(s[m.end():])
        return _NextWD(_wd(m.group(1)), ref), rest

    m = re.match(r"this:([A-Za-z]{3})", s, re.I)
    if m:
        ref, rest = _parse_from_clause(s[m.end():])
        return _ThisWD(_wd(m.group(1)), ref), rest

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
    s = s.strip()
    if _OFFSET_RE.match(s):
        node, rest = _Serve(), s
    else:
        node, rest = _parse_base(s)
    # offsets
    while True:
        m = _OFFSET_RE.match(rest)
        if not m:
            break
        node = _Offset(node, int(m.group(1)), m.group(2))
        rest = rest[m.end():]
    # optional trailing time-of-day suffix: @HH:MM or @HH:MM-HH:MM
    m = _TIME_RE.match(rest)
    if m:
        start = _parse_time(m.group(1), m.group(2))
        end = _parse_time(m.group(3), m.group(4)) if m.group(3) else None
        _check_work_hours(start, end)
        node = _TimeAttach(node, start, end)
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
    if isinstance(v, TimeInterval):
        return "timeinterval"
    if isinstance(v, Interval):
        return "interval"
    return "date"


_TOKEN_RE = re.compile(r"\{([^{}]*)\}")
_EMIT_RE = re.compile(r"^\s*!\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.+)$")


def render_body(body: str, ctx: Context) -> RenderResult:
    """Render every {token} in an email body to a concrete date string, and
    collect any {!name = expr} emissions into the anchor table.

    Emissions are evaluated against ctx (this email's serve date) and added to a
    local copy of the anchor table as we go, so a later token in the same body
    may reference an anchor an earlier token emitted.
    """
    emissions: dict[str, Value] = {}
    local = Context(ctx.serve, dict(ctx.anchors))

    def _sub(match: re.Match) -> str:
        inner = match.group(1).strip()
        emit = _EMIT_RE.match(inner)
        if emit:
            name, expr = emit.group(1), emit.group(2)
            value = resolve(expr, local)
            emissions[name] = value
            local.anchors[name] = value
            return human(value)
        return human(resolve(inner, local))

    text = _TOKEN_RE.sub(_sub, body)
    return RenderResult(text=text, emissions=emissions)


def _ordinal(n: int) -> str:
    if 10 <= n % 100 <= 20:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


def _clock(dt: datetime) -> str:
    h = dt.hour % 12 or 12
    ampm = "AM" if dt.hour < 12 else "PM"
    mm = "" if dt.minute == 0 else f":{dt.minute:02d}"
    return f"{h}{mm} {ampm}"


def _full_day(d) -> str:
    return f"{d.strftime('%A, %B')} {_ordinal(d.day)}, {d.year}"


def human(value: Value) -> str:
    """Natural-language rendering used in email bodies."""
    if isinstance(value, Interval):
        return f"the week of {value.start.strftime('%B')} {_ordinal(value.start.day)}, {value.start.year}"
    if isinstance(value, TimeInterval):
        return f"{_full_day(value.start)}, {_clock(value.start)}–{_clock(value.end)}"
    if isinstance(value, datetime):
        return f"{_full_day(value)} at {_clock(value)}"
    return _full_day(value)
