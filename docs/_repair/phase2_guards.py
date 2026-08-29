# PROTOTYPE — not shipped, not imported by sb/. Preserved for handoff.
#
# The phase-2 grader contract and its adversarial guard set. `grade_v4` is
# iteration 3 (2026-08-29), the version under verification; `grade_v3` is kept
# so the two can be compared. Read docs/grader-contract.md first.
#
# Run:
#   .venv/bin/python -c "
#   import sys; sys.path.insert(0,'docs/_repair')
#   from phase2_guards import *
#   from sb.grader import grade_email
#   report([run_guards('shipped', grade_email),
#           run_guards_named('v3', grade_v3),
#           run_guards_named('v4', grade_v4)])"
#
# run_guards_named runs the engine oracle with sb/oracle.py's title policy changed
# to op.name (blocker 3); run_guards uses sb/oracle.py as shipped. The register
# says this must become a real test suite in sb/tests/.
"""
Phase 2 — stronger guards first, then the contract.

Iteration 1 passed three guards that turned out to be vacuous. This guard set is
built from the adversaries the verifier used to break it, so a contract cannot
look safe by accident. Iteration 3 adds `dupmove` after finding the synth agent
itself double-booked on every move.

GUARDS (all free, all offline, all against the 1c capture's corpus/plan):
  real            certified claude-sonnet-4-5 capture      informational
  oracle_engine   sb/oracle.py via sb.engine.run           MUST be 167  (sb.scale gate)
  oracle_name     perfect agent, titles = obligation name  MUST be 167
  oracle_subject  perfect agent, titles = email subject    should be high
  oracle_inflect  perfect agent, inflected titles          robustness
  null            does nothing                             MUST be 64
  dup5            right answer + 5 duplicates of each      MUST NOT exceed null
  shot7/45/90     date-blind, N-day shotgun                MUST NOT exceed null
  wrongkind       right title/date, wrong kind             informational
  dupmove         right answer, but every move leaves the  MUST NOT exceed oracle_name - (move emails)
                  old object behind (double-booking)
"""
import json, glob, re
from datetime import date, timedelta
from pathlib import Path

import sb.grader as G
from sb.grader import grade_email
from sb.live.runner import _node_state, _turn_delta
from sb.oracle import _as_dt, _target, oracle_model
from sb.engine import Store, run as engine_run
from sb.resolver import Context
from sb.schema import load_corpus
from sb.scheduler import Levers, build_plan

CAP = 'captures/baseline-sonnet-4-5'
MAN = json.load(open(f'{CAP}/manifest.json'))
CORPUS = load_corpus(MAN['corpus_dir'])
LV = MAN['levers']
PLAN = build_plan(CORPUS, start_date=date.fromisoformat(MAN['start']), seed=MAN['seed'],
                  n_days=MAN['n_days'], levers=Levers(LV['daily_min'], LV['daily_max'], LV['urgency_horizon']))
DAYS = [json.loads(Path(p).read_text()) for p in sorted(glob.glob(f'{CAP}/days/*.json'))]

STOP = set("""the a an of to for and or with on in at by is are be this that your you our their its it
can could should would will may might must do does did have has had not no yes please need want make
made create creates created contact inform pick fill decide schedule set send get put let know talk
about discuss discussion meeting meet sync call review session event todo task reminder note update
follow plan planning arrange organize confirm check ensure prepare added new final go day date time
week weeks month list""".split())

_w = lambda s: re.findall(r'[a-z0-9]+', s.lower())


def stem(w):
    """Light suffix stripping so 'meetings' matches 'meeting', 'delayed' matches 'delay'."""
    for suf in ('ings', 'ing', 'ies', 'ed', 'es', 's', 'ly'):
        if len(w) > len(suf) + 2 and w.endswith(suf):
            return w[:-len(suf)]
    return w


def stems(s):
    return {stem(w) for w in _w(s)}


def keywords_of(op):
    ws = [w for w in _w(op.name.replace('_', ' ')) if w not in STOP]
    return {stem(w) for w in (ws or _w(op.name.replace('_', ' ')))}


def overlap(obj, kws):
    hay = stems(f"{obj.title} {obj.description}")
    return sum(k in hay for k in kws) / len(kws) if kws else 0.0


def overlap_title(obj, kws):
    hay = stems(obj.title)
    return sum(k in hay for k in kws) / len(kws) if kws else 0.0


# ------------------------------------------------------------------ contract
def grade_v3(answer, ctx, state, turn, eid=None, *, dateok_tiebreak=False,
             volume_brake=True, cross_kind=False):
    if not answer.ops:
        created = turn.events + turn.todos
        p = not created
        return G.EmailResult(passed=p, max=1, headline="no-action",
                             details=[{"passed": p, "label": "no-action", "expected": "take no action",
                                       "actual": "; ".join(G._fmt_obj(o) for o in created) or "(nothing)",
                                       "reason": "correctly took no action" if p else "over-acted"}])

    def pool_for(op):
        if not cross_kind:
            return state.events if op.kind == 'event' else state.todos
        return state.events + state.todos

    pairs = []
    for i, op in enumerate(answer.ops):
        if op.verb == 'cancel':
            continue
        kws = keywords_of(op)
        for o in pool_for(op):
            sc = overlap(o, kws)
            if sc <= 0:
                continue
            mine = 1 if (eid and o.email_id == eid) else 0
            kindok = 1 if (o.kind == op.kind) else 0
            key = (sc, kindok, mine, 1 if (dateok_tiebreak and G._predicate_ok(o, op.on, ctx, op.tolerance)) else 0)
            pairs.append((key, i, o))
    pairs.sort(key=lambda t: (-t[0][0], -t[0][1], -t[0][2], -t[0][3]))
    claimed_op, claimed_obj = {}, set()
    for key, i, o in pairs:
        if i in claimed_op or id(o) in claimed_obj:
            continue
        claimed_op[i] = (o, key[0])
        claimed_obj.add(id(o))

    details = []
    for i, op in enumerate(answer.ops):
        word = "event" if op.kind == "event" else "to-do"
        kw = op.name.replace('_', ' ')
        exp = f'{word} ~"{kw}" @ {G._describe_predicate(op.on, ctx)}'
        if op.verb == 'cancel':
            kws = keywords_of(op)
            left = [o for o in pool_for(op) if overlap(o, kws) >= 1.0]
            p = not left
            details.append({"passed": p, "label": f"cancel ~{kw}", "expected": f'{word} ~"{kw}" cancelled',
                            "actual": "; ".join(G._fmt_obj(o) for o in left) or "(nothing — cancelled)",
                            "reason": "cancelled" if p else f"should be cancelled, but {len(left)} still on the calendar"})
            continue
        got = claimed_op.get(i)
        if got is None:
            details.append({"passed": False, "label": f"{op.verb} ~{kw}", "expected": exp,
                            "actual": "(nothing matching created)",
                            "reason": f'no {word} titled like "{kw}" was created'})
            continue
        obj, sc = got
        # VOLUME BRAKE: an equally-good candidate this same email created, that no
        # obligation claimed, is a duplicate. Distinct obligations sharing vocabulary
        # are unaffected -- their objects are claimed by their own obligation.
        dups = []
        if volume_brake:
            kws = keywords_of(op)
            # Scope to objects created THIS TURN. Available on every code path, and
            # cannot be defeated by a mis-stamped email_id (register A-5). An extra
            # equally-matching object created now is a duplicate; one inherited from
            # an earlier email is that email's obligation, not an over-creation.
            # _turn_delta builds fresh Obj instances, so compare by value not identity.
            key = lambda o: (o.kind, o.title, o.when, o.email_id)
            fresh = {key(o) for o in (turn.events + turn.todos)}
            for o in pool_for(op):
                if id(o) in claimed_obj or o is obj or key(o) not in fresh:
                    continue
                if overlap(o, kws) >= sc:
                    dups.append(o)
        if dups:
            details.append({"passed": False, "label": f"{op.verb} ~{kw}", "expected": exp,
                            "actual": "; ".join(G._fmt_obj(o) for o in [obj] + dups[:3]),
                            "reason": f"over-created: {len(dups)+1} equally-matching {word}s for one obligation"})
            continue
        ok = G._predicate_ok(obj, op.on, ctx, op.tolerance)
        details.append({"passed": ok, "label": f"{op.verb} ~{kw}", "expected": exp,
                        "actual": G._fmt_obj(obj), "reason": "matched" if ok else "on the wrong day"})
    return G.EmailResult(passed=all(d["passed"] for d in details), max=1, details=details,
                         headline="; ".join(d["reason"] for d in details))


# -------------------------------------------------------------- synth agents
def synth(policy, shift=0, act=True, dup=0, shotgun=0, wrongkind=False, dupmove=False):
    """A scheduling-perfect agent with a pluggable title policy.

    It remembers which store records it made for each obligation, so a move or a
    cancel acts on the right object regardless of how it was titled -- exactly what
    a real agent does with the ids the store hands back. `dupmove=True` makes it
    move by creating a copy and leaving the old one behind (the double-booking the
    prompt forbids), which is the `dupmove` guard.
    """
    ev, td, out, n = [], [], [], 0
    owned = {}                                   # (node, obligation) -> [record ids]
    for day_no, batch in enumerate([b for b in PLAN.per_day if b], 1):
        before = {r['id'] for r in ev + td}
        if act:
            for eid in batch:
                em = CORPUS.emails[eid]
                ctx = Context(serve=PLAN.serve_date[eid], anchors=PLAN.anchors)
                for op in em.answer.ops:
                    base = op.name.replace('_', ' ')
                    if policy == 'subject':   title = em.subject
                    elif policy == 'inflect': title = " ".join(w + 's' for w in base.split())
                    else:                     title = base
                    if op.verb == 'cancel' or (op.verb == 'move' and not dupmove):
                        gone = set(owned.pop((em.node, op.name), []))
                        ev[:] = [r for r in ev if r['id'] not in gone]
                        td[:] = [r for r in td if r['id'] not in gone]
                        if op.verb == 'cancel':
                            continue
                    kind = op.kind
                    if wrongkind: kind = 'todo' if op.kind == 'event' else 'event'
                    when0 = _as_dt(_target(op.on, ctx), 9 if kind == 'event' else 17)
                    spread = range(shotgun) if shotgun else [0]
                    made = []
                    for d in spread:
                        when = when0 + timedelta(days=shift)
                        if shotgun:
                            when = _as_dt(PLAN.serve_date[eid], 9) + timedelta(days=d)
                        for c in range(1 + dup):
                            n += 1
                            rec = {'id': f'x_{n}', 'email_id': eid, 'title': title, 'description': ''}
                            if kind == 'event':
                                rec['start'] = when.isoformat(); rec['end'] = (when + timedelta(hours=1)).isoformat(); ev.append(rec)
                            else:
                                rec['due_date'] = when.isoformat(); td.append(rec)
                            made.append(rec['id'])
                    owned.setdefault((em.node, op.name), []).extend(made)
        out.append({'day': day_no, 'batch': list(batch), 'ok': True,
                    'state': {'events': list(ev), 'todos': list(td)},
                    'day_new': sorted({r['id'] for r in ev + td} - before)})
    return out


class _O:
    def __init__(self, rec, kind):
        self.title = rec.get('title', ''); self.description = rec.get('description', ''); self.kind = kind


def score(day_records, grader, **kw):
    ep = 0
    for rec in day_records:
        st, dn = rec['state'], set(rec['day_new'])
        by = {r['id']: r.get('email_id', '') for r in st['events'] + st['todos']}
        for eid in rec['batch']:
            em = CORPUS.emails[eid]
            ctx = Context(serve=PLAN.serve_date[eid], anchors=PLAN.anchors)
            nw = {i for i in dn if by.get(i) == eid}
            ns, tt = _node_state(CORPUS, st, em.node, nw), _turn_delta(CORPUS, st, nw)
            r = grader(em.answer, ctx, ns, tt, eid=eid, **kw) if grader is not grade_email else grade_email(em.answer, ctx, ns, tt)
            ep += bool(r.passed)
    return ep


def oracle_engine_score(grader, **kw):
    """sb/oracle.py through sb.engine.run — the repo's own mandatory gate."""
    if grader is grade_email:
        return engine_run(CORPUS, PLAN, oracle_model, store=Store(CORPUS)).passed
    orig = G.grade_email
    G.grade_email = lambda a, c, s, t: grader(a, c, s, t, eid=None, **kw)
    try:
        import sb.engine as E
        E.grade_email = G.grade_email
        return engine_run(CORPUS, PLAN, oracle_model, store=Store(CORPUS)).passed
    finally:
        G.grade_email = orig
        import sb.engine as E
        E.grade_email = orig


WORLDS = [
    ('real',           lambda: DAYS),
    ('oracle_name',    lambda: synth('name')),
    ('oracle_subject', lambda: synth('subject')),
    ('oracle_inflect', lambda: synth('inflect')),
    ('null',           lambda: synth('name', act=False)),
    ('dup5',           lambda: synth('name', dup=5)),
    ('shot7',          lambda: synth('name', shotgun=7)),
    ('shot45',         lambda: synth('name', shotgun=45)),
    ('shot90',         lambda: synth('name', shotgun=90)),
    ('wrongkind',      lambda: synth('name', wrongkind=True)),
    ('dupmove',        lambda: synth('name', dupmove=True)),
]


def run_guards(label, grader, **kw):
    row = {}
    for name, mk in WORLDS:
        row[name] = score(mk(), grader, **kw)
    row['oracle_engine'] = oracle_engine_score(grader, **kw)
    return label, row


N_MOVE_EMAILS = sum(1 for e in CORPUS.emails.values() if any(op.verb == 'move' for op in e.answer.ops))


def report(rows):
    cols = ['real', 'oracle_engine', 'oracle_name', 'oracle_subject', 'oracle_inflect',
            'null', 'dup5', 'shot7', 'shot45', 'shot90', 'wrongkind', 'dupmove']
    print(f"{'contract':26s} " + " ".join(f"{c[:9]:>9s}" for c in cols))
    print("-" * (27 + 10 * len(cols)))
    for label, r in rows:
        print(f"{label:26s} " + " ".join(f"{r.get(c,0):9d}" for c in cols))
    print("\nGUARDS: oracle_engine & oracle_name MUST be 167 · null MUST be 64")
    print("        dup5 / shot7 / shot45 / shot90 MUST NOT exceed null (64)")
    print(f"        dupmove MUST NOT exceed oracle_name - {N_MOVE_EMAILS} (every move email double-booked)")
    for label, r in rows:
        bad = []
        if r.get('oracle_engine') != 167: bad.append(f"oracle_engine {r['oracle_engine']}")
        if r.get('oracle_name') != 167:   bad.append(f"oracle_name {r['oracle_name']}")
        if r.get('null') != 64:           bad.append(f"null {r['null']}")
        for g in ('dup5', 'shot7', 'shot45', 'shot90'):
            if r.get(g, 0) > r.get('null', 64): bad.append(f"{g} {r[g]}")
        if r.get('dupmove', 0) > r.get('oracle_name', 167) - N_MOVE_EMAILS: bad.append(f"dupmove {r['dupmove']}")
        print(f"  {label:26s} {'PASS' if not bad else 'FAIL -> ' + ', '.join(bad)}")


# ------------------------------------------------------------ contract, iter 3
# grade_v4 closes the four blockers listed in docs/grader-contract.md:
#   kind      the pool is cross-kind; kind is the FIRST sort key so a right-kind object
#             always wins, and a wrong-kind claim fails with its own reason instead of
#             reading as "nothing created" (G-9). Score-neutral by construction vs the
#             kind-filtered pool except where cancel/brake now SEE the object.
#   cancel    graded THROUGH the assignment: create/move claim first, then a cancel
#             claims the best unclaimed object at overlap >= cancel_tau and fails if it
#             gets one. A sibling op's correctly-kept object can no longer fail a cancel.
#   mine      derived from turn membership (by value), not email_id. The engine path
#             passes no email id, so v3's tie-break was dead there and the two grading
#             paths disagreed on the same state (found while closing blocker 4).
#   move      a `move` additionally fails on a STALE SURVIVOR: an unclaimed same-kind
#             object from any earlier turn matching at >= the claimed score. A move's
#             obligation already had an object; after the move exactly one may remain.
def _vkey(o):
    return (o.kind, o.title, o.when, o.email_id)


def grade_v4(answer, ctx, state, turn, eid=None, *, cancel_tau=1.0, move_stale=True,
             kind_visible=True, kind_first=True, brake_hay='full'):
    if not answer.ops:
        created = turn.events + turn.todos
        p = not created
        return G.EmailResult(passed=p, max=1, headline="no-action",
                             details=[{"passed": p, "label": "no-action", "expected": "take no action",
                                       "actual": "; ".join(G._fmt_obj(o) for o in created) or "(nothing)",
                                       "reason": "correctly took no action" if p else "over-acted"}])

    fresh = {_vkey(o) for o in (turn.events + turn.todos)}
    is_fresh = lambda o: _vkey(o) in fresh

    def pool_for(op):
        if kind_visible:
            return state.events + state.todos
        return state.events if op.kind == 'event' else state.todos

    def rank(op, o, sc):
        kindok = 1 if o.kind == op.kind else 0
        return (kindok, sc, 1 if is_fresh(o) else 0) if kind_first else (sc, kindok, 1 if is_fresh(o) else 0)

    kws = {i: keywords_of(op) for i, op in enumerate(answer.ops)}
    claimed_op, claimed_obj = {}, set()

    def assign(verbs, floor):
        pairs = []
        for i, op in enumerate(answer.ops):
            if op.verb not in verbs:
                continue
            for o in pool_for(op):
                sc = overlap(o, kws[i])
                if sc <= 0 or sc < floor:
                    continue
                pairs.append((rank(op, o, sc), i, o, sc))
        pairs.sort(key=lambda t: t[0], reverse=True)
        for _, i, o, sc in pairs:
            if i in claimed_op or id(o) in claimed_obj:
                continue
            claimed_op[i] = (o, sc)
            claimed_obj.add(id(o))

    assign(('create', 'move'), 0.0)          # phase 1: the work this email asked for
    assign(('cancel',), cancel_tau)          # phase 2: cancels take what is left over

    details = []
    for i, op in enumerate(answer.ops):
        word = "event" if op.kind == "event" else "to-do"
        kw = op.name.replace('_', ' ')
        got = claimed_op.get(i)
        if op.verb == 'cancel':
            p = got is None
            details.append({"passed": p, "label": f"cancel ~{kw}", "expected": f'{word} ~"{kw}" cancelled',
                            "actual": G._fmt_obj(got[0]) if got else "(nothing — cancelled)",
                            "reason": "cancelled" if p else "should be cancelled, but still on the calendar"})
            continue
        exp = f'{word} ~"{kw}" @ {G._describe_predicate(op.on, ctx)}'
        if got is None:
            details.append({"passed": False, "label": f"{op.verb} ~{kw}", "expected": exp,
                            "actual": "(nothing matching created)",
                            "reason": f'no {word} titled like "{kw}" was created'})
            continue
        obj, sc = got
        if obj.kind != op.kind:
            details.append({"passed": False, "label": f"{op.verb} ~{kw}", "expected": exp,
                            "actual": G._fmt_obj(obj),
                            "reason": f"wrong kind: created a {'event' if obj.kind == 'event' else 'to-do'}, expected a {word}"})
            continue
        # Over-creation is judged on TITLES only. Identity may use the description
        # (a real model puts the obligation there, verified flips in VERIFY-phase2 §5),
        # but a sibling whose description merely mentions this obligation is not a
        # copy of it (G-3). Floor: the claimed object's own title score, and > 0.
        dups, stale = [], []
        hay = (lambda o: overlap(o, kws[i])) if brake_hay == 'full' else (lambda o: overlap_title(o, kws[i]))
        floor = sc if brake_hay == 'full' else hay(obj)
        for o in pool_for(op):
            if o.kind != op.kind or id(o) in claimed_obj or o is obj:
                continue
            h = hay(o)
            if h <= 0 or h < floor:
                continue
            if is_fresh(o):
                dups.append(o)
            elif op.verb == 'move' and move_stale:
                stale.append(o)
        if dups:
            details.append({"passed": False, "label": f"{op.verb} ~{kw}", "expected": exp,
                            "actual": "; ".join(G._fmt_obj(o) for o in [obj] + dups[:3]),
                            "reason": f"over-created: {len(dups)+1} equally-matching {word}s for one obligation"})
            continue
        if stale:
            details.append({"passed": False, "label": f"{op.verb} ~{kw}", "expected": exp,
                            "actual": "; ".join(G._fmt_obj(o) for o in [obj] + stale[:3]),
                            "reason": f"moved, but {len(stale)} stale copy left behind (double-booked)"})
            continue
        ok = G._predicate_ok(obj, op.on, ctx, op.tolerance)
        details.append({"passed": ok, "label": f"{op.verb} ~{kw}", "expected": exp,
                        "actual": G._fmt_obj(obj), "reason": "matched" if ok else "on the wrong day"})
    passed = all(d["passed"] for d in details)
    headline = "; ".join(d["expected"] for d in details) if passed else next(d["reason"] for d in details if not d["passed"])
    return G.EmailResult(passed=passed, max=1, headline=headline, details=details)


def oracle_model_name(email, rendered_body, ctx, store):
    """sb/oracle.py with blocker 3 applied: titles by op.name, not op.match."""
    from sb.oracle import _create
    for op in email.answer.ops:
        title = op.name.replace('_', ' ')
        if op.verb == 'cancel':
            while (oid := store.find_in_node(email.node, op.kind, title)) is not None:
                store.delete(oid)
            continue
        when = _as_dt(_target(op.on, ctx), 9 if op.kind == 'event' else 17)
        existing = store.find_in_node(email.node, op.kind, title)
        if op.verb == 'move':
            if existing is not None and op.kind == 'event':
                store.update_event(existing, start=when)
            elif existing is not None:
                store.delete(existing)
                store.create_todo(email.id, title, when)
            else:
                _create(store, email.id, op.kind, title, when)
        else:
            _create(store, email.id, op.kind, title, when)


def oracle_engine_score_named(grader, **kw):
    """oracle_engine with the op.name-titled oracle (what sb/oracle.py becomes)."""
    orig = G.grade_email
    import sb.engine as E
    G.grade_email = E.grade_email = (grade_email if grader is grade_email
                                     else (lambda a, c, s, t: grader(a, c, s, t, eid=None, **kw)))
    try:
        return engine_run(CORPUS, PLAN, oracle_model_name, store=Store(CORPUS)).passed
    finally:
        G.grade_email = E.grade_email = orig


def run_guards_named(label, grader, **kw):
    """run_guards, but oracle_engine uses the op.name-titled oracle."""
    row = {}
    for name, mk in WORLDS:
        row[name] = score(mk(), grader, **kw)
    row['oracle_engine'] = oracle_engine_score_named(grader, **kw)
    return label, row
