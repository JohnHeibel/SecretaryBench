"""
Corpus-wide probes backing the phase 1d hand-grade findings.

Every number in docs/_repair/handgrade-findings.md comes from here. Read-only:
touches no corpus, capture or code, and does not move the corpus hash.

    .venv/bin/python docs/_repair/handgrade_probes.py

Three probes:
  1. kind convention      -- how the corpus keys actions vs gatherings (K-7 evidence)
  2. hedged-time control  -- falsifies "unconfirmed time implies todo"
  3. title recoverability -- are a key's identity words even present in the email?
                             (the title analogue of K-6; no register ID yet)
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from sb.grader import keywords_of          # noqa: E402
from sb.schema import load_corpus          # noqa: E402

GATHER = re.compile(r"conversation|meeting|call|1:1|one.on.one|sync|dinner|lunch|interview"
                    r"|chat|briefing|showcase|mixer|webinar|standup|walk.?through", re.I)
VERBS = {"ask", "send", "review", "approve", "approval", "sign", "order", "decide", "decision",
         "confirm", "contact", "create", "fill", "talk", "thank", "add", "double", "check",
         "draft", "prepare", "submit", "push"}
HEDGE = re.compile(r"\b(i'?m thinking|i think|maybe|perhaps|probably|tentativ|not sure|pencil"
                   r"|penciled|proposed|propose|let'?s say|how about|if that works|hopefully"
                   r"|does .{0,15}work"
                   r"|might be|leaning|prelim)", re.I)


def creates(corpus):
    emails = corpus.emails.values() if isinstance(corpus.emails, dict) else corpus.emails
    for e in emails:
        if not e.answer:
            continue
        for op in e.answer.ops:
            if op.verb in ("create", "move"):
                yield e, op


def first_tok(name: str) -> str:
    parts = re.split(r"[\s_\-]+", name.strip().lower())
    return parts[0] if parts else ""


def toks(s: str) -> set[str]:
    return {t for t in re.split(r"[^a-z0-9]+", (s or "").lower()) if t}


def present(word: str, have: set[str]) -> bool:
    """Generous stem test, so a flag means the word is really absent."""
    return any(h.startswith(word[:4]) or word.startswith(h[:4]) for h in have if len(h) > 2)


def main() -> None:
    c = load_corpus("corpus")
    ops = list(creates(c))

    # --- probe 1: kind convention -------------------------------------------------
    verb_named = {"event": [], "todo": []}
    gathering = {"event": [], "todo": []}
    for e, op in ops:
        if first_tok(op.name) in VERBS:
            verb_named[op.kind].append((op.name.strip(), e.id))
        if GATHER.search(op.name):
            gathering[op.kind].append((op.name.strip(), e.id))

    print("=" * 72)
    print("PROBE 1  kind convention")
    print(f"  verb-named ops : todo {len(verb_named['todo']):>3} | event {len(verb_named['event']):>3}  <- event = outlier")
    for n, i in verb_named["event"]:
        print(f"        {n!r:34} {i}")
    print(f"  gathering nouns: event {len(gathering['event']):>3} | todo  {len(gathering['todo']):>3}  <- todo = outlier")
    for n, i in gathering["todo"]:
        print(f"        {n!r:34} {i}")

    # --- probe 2: hedged-time control ---------------------------------------------
    hedged_events = [(op.name.strip(), HEDGE.search(e.body).group(0), e.id)
                     for e, op in ops
                     if op.kind == "event" and op.verb == "create" and HEDGE.search(e.body or "")]
    print("=" * 72)
    print("PROBE 2  hedged-time control (falsifies 'unconfirmed time implies todo')")
    print(f"  event-keyed creates on a hedged email: {len(hedged_events)}")
    for n, h, i in hedged_events:
        print(f"        {n!r:34} hedge={h!r:20} {i}")

    # --- probe 3: title recoverability --------------------------------------------
    # Conservative: reads the RAW body, which still carries anchor token names the
    # model never sees (a rendered date replaces them). So this OVER-states what is
    # available to the model and the counts below are a FLOOR on the defect.
    none_present, some_absent = [], []
    for e, op in ops:
        kw = keywords_of(op)
        if not kw:
            continue
        have = toks(e.body) | toks(e.subject)
        missing = sorted(w for w in kw if not present(w, have))
        if len(missing) == len(kw):
            none_present.append((op.name.strip(), sorted(kw), e.id))
        elif missing:
            some_absent.append((op.name.strip(), sorted(kw), missing, e.id))

    print("=" * 72)
    print("PROBE 3  title recoverability  (title analogue of K-6; no register ID yet)")
    print(f"  create/move ops                    : {len(ops)}")
    print(f"  EVERY identity word absent         : {len(none_present)}")
    for n, k, i in none_present:
        print(f"        {n!r:34} kw={k} | {i}")
    print(f"  at least one identity word absent  : {len(some_absent)}")
    for n, k, m, i in some_absent:
        print(f"        {n!r:34} kw={k} missing={m} | {i}")
    print("=" * 72)


if __name__ == "__main__":
    main()
