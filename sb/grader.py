"""
sb.grader — the state-based grader.

Grades emails against their answer keys by inspecting the calendar/todo STATE
(not the model's edits), so a reschedule done as update-in-place and one done as
delete-then-recreate score identically (ANSWER_KEY_GRAMMAR.md §8).

The answer key is a list of `ops` (create / move / cancel) on named obligations.
The identity contract (docs/grader-contract.md) reconciles them against the
cumulative node state:

  - identity    : an object belongs to an obligation when the stemmed content words
                  of the obligation's NAME appear in `title + description`. The
                  fraction present is the pair's score. `op.match` is not consulted
                  (register G-1, G-8, C-10).
  - assignment  : every op of the email claims at most one same-kind object and
                  every object serves at most one op. Best title match wins, then
                  best title + description match, then an object created today,
                  then the object whose title is most about the obligation, then
                  create/move before cancel. A create may claim only an object
                  created today; a move may claim any; a cancel enters only at
                  full overlap (register G-2, G-7).
  - create/move : the claimed object must be on the right day, and no equally-
                  matching same-kind object created today may be left unclaimed —
                  that is a duplicate, whichever email of the node it is stamped
                  to. A move additionally fails if an unclaimed same-kind object
                  from an earlier day still carries more than half the
                  obligation's words — the stale copy of a double-booking.
  - cancel      : fails if it claims anything, i.e. if a same-kind object carrying
                  every content word of the obligation survives that no other op
                  of the email accounts for.
  - no ops      : a no-action / FYI / bait email — its turn must create nothing.

Inputs per email:
  - answer:   the Answer (ops)
  - ctx:      resolver.Context (this email's serve date + the live anchor table)
  - state:    NodeState — cumulative objects belonging to this email's NODE
  - turn:     TurnDelta — objects CREATED for this email (its email_id stamp, or
              its own model call on the engine path)
  - today:    TurnDelta — everything created today, whoever it is stamped to
              (defaults to `turn`, which is exact on the engine path)

Output: EmailResult(passed, max=1, details[]). Binary per email.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Optional

from sb import resolver
from sb.resolver import Context, Interval, Value
from sb.schema import Answer, Op


@dataclass
class Obj:
    """A calendar event or todo in the store (normalized for grading)."""
    kind: str                 # "event" | "todo"
    title: str
    when: datetime            # event.start or todo.due_date
    email_id: str             # attribution tag the model passed
    description: str = ""
    end: Optional[datetime] = None    # event end (for duration checks)


@dataclass
class NodeState:
    events: list[Obj] = field(default_factory=list)
    todos: list[Obj] = field(default_factory=list)


@dataclass
class TurnDelta:
    events: list[Obj] = field(default_factory=list)
    todos: list[Obj] = field(default_factory=list)


@dataclass
class EmailResult:
    passed: bool
    max: int
    details: list[dict] = field(default_factory=list)
    headline: str = ""


# --- identity: which object is which obligation ----------------------------

# Words that carry no identity in an obligation name. Matched on the raw word.
# Measured insensitive: removing any one moves no guard by more than 2.
STOP_WORDS = frozenset("""the a an of to for and or with on in at by is are be this that your you our their its it
can could should would will may might must do does did have has had not no yes please need want make
made create creates created contact inform pick fill decide schedule set send get put let know talk
about discuss discussion meeting meet sync call review session event todo task reminder note update
follow plan planning arrange organize confirm check ensure prepare added new final go day date time
week weeks month list""".split())

_WORD_RE = re.compile(r"[a-z0-9]+")


def _words(s: str) -> list[str]:
    return _WORD_RE.findall(s.lower())


def _stem(w: str) -> str:
    """Light suffix stripping so 'meetings' matches 'meeting', 'trophies' matches
    'trophy', 'delayed' matches 'delay'. -es only after a sibilant, so 'notes'
    still meets 'note'; -s not after another s, so 'press' stays whole."""
    if len(w) > 5 and w.endswith("ies"):
        return w[:-3] + "y"
    for suf in ("ings", "ing", "ed", "es", "s", "ly"):
        if len(w) > len(suf) + 2 and w.endswith(suf):
            if suf == "es" and not w[:-2].endswith(("s", "x", "z", "ch", "sh")):
                continue
            if suf == "s" and w[-2] == "s":
                continue
            return w[: -len(suf)]
    return w


_STOP_STEMS = frozenset(_stem(w) for w in STOP_WORDS)


def _stems(s: str) -> set[str]:
    return {_stem(w) for w in _words(s)}


def keywords_of(op: Op) -> set[str]:
    """The stemmed content words of the obligation's name. A name made only of
    stop-words (e.g. `event day!`) falls back to all of its words."""
    words = _words(op.name.replace("_", " "))
    content = [w for w in words if w not in STOP_WORDS]
    return {_stem(w) for w in (content or words)}


def overlap(obj: Obj, keywords: set[str]) -> float:
    """Fraction of the obligation's keywords present in the object's title + description."""
    if not keywords:
        return 0.0
    hay = _stems(f"{obj.title} {obj.description}")
    return sum(k in hay for k in keywords) / len(keywords)


def _title_overlap(obj: Obj, keywords: set[str]) -> float:
    """The same fraction over the title alone: the primary evidence. The
    description can complete a match but never outranks a title match (G-3)."""
    if not keywords:
        return 0.0
    hay = _stems(obj.title)
    return sum(k in hay for k in keywords) / len(keywords)


def _title_precision(obj: Obj, keywords: set[str]) -> float:
    """Share of the title's content words that are the obligation's keywords. A
    tie-break only: among equally-matching objects prefer the one that is ABOUT
    the obligation over one that merely mentions it, or that matched through
    its description alone (which scores 0 here)."""
    ts = {w for w in _stems(obj.title) if w not in _STOP_STEMS} or _stems(obj.title)
    return sum(k in ts for k in keywords) / len(ts) if ts else 0.0


# --- date predicates --------------------------------------------------------

def _to_date(v: Value) -> date:
    return v.date() if isinstance(v, datetime) else (v.start if isinstance(v, Interval) else v)


def _matches_value(obj_when: datetime, expected: Value, tolerance: str) -> bool:
    if isinstance(expected, Interval):
        return expected.contains(obj_when)
    if tolerance.startswith("within:"):
        n = int(tolerance.split(":", 1)[1].rstrip("d"))
        return abs((obj_when.date() - _to_date(expected)).days) <= n
    # day equality (the benchmark grades at whole-day granularity)
    return obj_when.date() == _to_date(expected)


def _predicate_ok(obj: Obj, predicate: Optional[dict], ctx: Context, tolerance: str) -> bool:
    if not predicate:
        return True
    if "eq" in predicate:
        return _matches_value(obj.when, resolver.resolve(predicate["eq"], ctx), tolerance)
    if "any_of" in predicate:
        return any(_matches_value(obj.when, resolver.resolve(e, ctx), tolerance)
                   for e in predicate["any_of"])
    if "by" in predicate:
        return ctx.serve <= obj.when.date() <= _to_date(resolver.resolve(predicate["by"], ctx))
    if "in" in predicate:
        window = resolver.resolve(predicate["in"], ctx)
        ok = isinstance(window, Interval) and window.contains(obj.when)
        if ok and "not_in" in predicate:
            block = resolver.resolve(predicate["not_in"], ctx)
            if isinstance(block, Interval) and block.contains(obj.when):
                ok = False
        return ok
    return True


# --- human-readable formatting (so a run explains itself) ------------------

def _h12(dt: datetime) -> str:
    h = dt.hour % 12 or 12
    ap = "AM" if dt.hour < 12 else "PM"
    mm = "" if dt.minute == 0 else f":{dt.minute:02d}"
    return f"{h}{mm} {ap}"


def _fmt_value(v: Value) -> str:
    if isinstance(v, Interval):
        return f"{v.start.strftime('%a %b %d')}–{v.end.strftime('%b %d')}"
    if isinstance(v, datetime):
        return f"{v.strftime('%a %b %d')} {_h12(v)}"
    return v.strftime("%a %b %d")


def _fmt_obj(o: Obj) -> str:
    when = f"{o.when.strftime('%a %b %d')} {_h12(o.when)}"
    if o.kind == "event" and o.end is not None:
        when += f" ({round((o.end - o.when).total_seconds() / 60)}m)"
    return f'"{o.title}" {when}'


def _describe_predicate(predicate: Optional[dict], ctx: Context) -> str:
    if not predicate:
        return "(any time)"
    if "eq" in predicate:
        return _fmt_value(resolver.resolve(predicate["eq"], ctx))
    if "any_of" in predicate:
        return " or ".join(_fmt_value(resolver.resolve(e, ctx)) for e in predicate["any_of"])
    if "by" in predicate:
        return f"by {_fmt_value(resolver.resolve(predicate['by'], ctx))}"
    if "in" in predicate:
        s = f"within {_fmt_value(resolver.resolve(predicate['in'], ctx))}"
        if "not_in" in predicate:
            s += f" (avoiding {_fmt_value(resolver.resolve(predicate['not_in'], ctx))})"
        return s
    return "(any time)"


def _kind_word(kind: Optional[str]) -> str:
    return "event" if kind == "event" else "to-do"


# --- the contract -----------------------------------------------------------

CANCEL_OVERLAP = 1.0   # a cancel survivor must carry every content word of the obligation
HALF = 0.5             # "more than half the words": the stale-copy floor and the wrong-kind report floor


def _turn_key(o: Obj) -> tuple:
    # Turn membership is tested by value: the live runner rebuilds Obj instances
    # per call, so identity would never match there.
    return (o.kind, o.title, o.when, o.email_id)


def _no_action(turn: TurnDelta) -> EmailResult:
    created = turn.events + turn.todos
    passed = not created
    actual = "; ".join(_fmt_obj(o) for o in created) if created else "(nothing)"
    reason = "correctly took no action" if passed else f"over-acted — created {actual}"
    return EmailResult(passed=passed, max=1, headline=reason, details=[{
        "passed": passed, "label": "no-action",
        "expected": "take no action (FYI / no scheduling needed)",
        "actual": actual, "reason": reason}])


def grade_email(answer: Answer, ctx: Context, state: NodeState, turn: TurnDelta,
                today: Optional[TurnDelta] = None) -> EmailResult:
    """Grade one email against its node's state.

    `turn` is what was created for this email (its email_id stamp on the live
    path; its own model call on the engine path) and feeds the no-action rule.
    `today` is everything created today, whoever it is stamped to; it defaults
    to `turn`, which is exact on the engine path. The live runner and sb.regrade
    pass the whole day's new objects (see sb.live.runner._grade_day), so a
    duplicate cannot hide behind a sibling's stamp (register A-5).
    """
    if today is None:
        today = turn
    if not answer.ops:
        return _no_action(turn)

    pool = state.events + state.todos
    fresh = {_turn_key(o) for o in today.events + today.todos}
    keywords = {oi: keywords_of(op) for oi, op in enumerate(answer.ops)}

    # Exclusive greedy assignment over one ranking. The key is a total order on the
    # object's content and stamp, so the outcome cannot depend on the order the
    # store lists objects in; the stamp decides nothing unless two objects are
    # otherwise identical.
    pairs = []
    for oi, op in enumerate(answer.ops):
        kws = keywords[oi]
        for o in pool:
            if o.kind != op.kind:
                continue
            if op.verb == "create" and _turn_key(o) not in fresh:
                continue                       # a create asks for an object made now
            sc = overlap(o, kws)
            if sc <= 0 or (op.verb == "cancel" and sc < CANCEL_OVERLAP):
                continue
            pairs.append(((-_title_overlap(o, kws),                 # title evidence first
                           -sc,                                    # then title + description
                           0 if _turn_key(o) in fresh else 1,      # created today
                           -_title_precision(o, kws),              # title is about the obligation
                           1 if op.verb == "cancel" else 0,        # a perfect tie goes to the work
                           oi, o.title, o.description, o.when.isoformat(), o.email_id),
                          oi, o, sc))
    pairs.sort(key=lambda t: t[0])
    claimed: dict[int, tuple[Obj, float]] = {}
    taken: set[int] = set()
    for _, oi, o, sc in pairs:
        if oi in claimed or id(o) in taken:
            continue
        claimed[oi] = (o, sc)
        taken.add(id(o))

    details = []
    for oi, op in enumerate(answer.ops):
        word = _kind_word(op.kind)
        name = op.name.replace("_", " ")
        kws = keywords[oi]
        got = claimed.get(oi)

        if op.verb == "cancel":
            passed = got is None
            details.append({
                "passed": passed, "label": f"cancel ~{name}",
                "expected": f'{word} ~"{name}" cancelled',
                "actual": _fmt_obj(got[0]) if got else "(nothing — cancelled)",
                "reason": "cancelled" if passed else "should be cancelled, but still on the calendar"})
            continue

        expected = f'{word} ~"{name}" @ {_describe_predicate(op.on, ctx)}'
        label = f"{op.verb} ~{name}"
        if got is None:
            # Report, never claim, an object of the other kind that is clearly this
            # obligation: more than half its words, some of them in the title, not a
            # sibling's answer, and (for a create) made today.
            other = [(overlap(o, kws), o) for o in pool
                     if o.kind != op.kind and id(o) not in taken and _title_overlap(o, kws) > 0
                     and (op.verb != "create" or _turn_key(o) in fresh)]
            best_sc, best = max(other, key=lambda t: t[0], default=(0.0, None))
            if best is not None and best_sc > HALF:
                details.append({
                    "passed": False, "label": label, "expected": expected, "actual": _fmt_obj(best),
                    "reason": f"wrong kind: created a {_kind_word(best.kind)}, expected a {word}"})
            else:
                details.append({
                    "passed": False, "label": label, "expected": expected,
                    "actual": "(nothing matching created)",
                    "reason": f'no {word} titled like "{name}" was {"moved" if op.verb == "move" else "created"}'})
            continue

        obj, score = got
        dups, stale = [], []
        for o in pool:
            if o.kind != op.kind or id(o) in taken:
                continue
            h = overlap(o, kws)
            if _turn_key(o) in fresh:
                if h >= score:
                    dups.append(o)
            elif op.verb == "move" and h > HALF:
                stale.append(o)
        if dups:
            details.append({
                "passed": False, "label": label, "expected": expected,
                "actual": "; ".join(_fmt_obj(o) for o in [obj] + dups[:3]),
                "reason": f"over-created: {len(dups) + 1} equally-matching {word}s for one obligation"})
            continue
        if stale:
            details.append({
                "passed": False, "label": label, "expected": expected,
                "actual": "; ".join(_fmt_obj(o) for o in [obj] + stale[:3]),
                "reason": f"moved, but {len(stale)} stale copy left behind (double-booked)"})
            continue

        passed = _predicate_ok(obj, op.on, ctx, op.tolerance)
        details.append({
            "passed": passed, "label": label, "expected": expected, "actual": _fmt_obj(obj),
            "reason": "matched" if passed else "on the wrong day"})

    passed = all(d["passed"] for d in details)
    if passed:
        headline = "; ".join(d["expected"] for d in details)
    else:
        headline = next(d["reason"] for d in details if not d["passed"])
    return EmailResult(passed=passed, max=1, headline=headline, details=details)
