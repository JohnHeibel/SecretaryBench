"""
sb.schema — the corpus data model, JSON loader, and linter.

A corpus is a set of node files (corpus/nodes/*.json). A node groups emails that
share a cast and emitted anchors, but all scheduling is computed from a single
FLAT email-level DAG (see SERVING_AND_SCHEMA.md §0). This module loads the files,
flattens node-level sugar into email edges, derives the anchor emission map, and
lints the result (acyclic, anchors reachable, every token parses).
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from sb import resolver

_ANCHOR_REF_RE = re.compile(r"@([A-Za-z_][A-Za-z0-9_]*)")
_EMIT_IN_BODY_RE = re.compile(r"\{\s*!\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*([^{}]*)\}")
_TOKEN_RE = re.compile(r"\{([^{}]*)\}")

VALID_ACTIONS = {"create_event", "create_todo", "reschedule", "reply", "delegate"}
EDGE_TYPES = {"static", "date"}


class CorpusError(ValueError):
    """Raised when a corpus fails to load or lint."""


@dataclass
class Edge:
    email: str            # prerequisite email id
    type: str             # "static" | "date"


@dataclass
class ExpectEntry:
    action: str
    title_match: list[str] = field(default_factory=list)
    start: Optional[dict] = None          # predicate: {"eq"|"in"|"by"|"any_of"|"not_in": expr}
    due: Optional[dict] = None            # predicate (todos)
    count: Optional[int] = None
    tolerance: str = "exact_day"          # exact_day | within:Nd  (whole-day granularity)
    recurrence: Optional[dict] = None


@dataclass
class ForbidEntry:
    action: Optional[str] = None          # None = any action
    title_match: list[str] = field(default_factory=list)


@dataclass
class Answer:
    expect: list[ExpectEntry] = field(default_factory=list)
    forbid: list[ForbidEntry] = field(default_factory=list)
    emits: dict[str, str] = field(default_factory=dict)   # name -> expr (explicit, non-rendered)


@dataclass
class Email:
    id: str
    node: str
    sender: str
    recipients: list[str]
    subject: str
    body: str
    depends_on: list[Edge] = field(default_factory=list)
    answer: Answer = field(default_factory=Answer)

    # derived
    emits: dict[str, str] = field(default_factory=dict)    # name -> expr (body + answer.emits)
    anchor_refs: set[str] = field(default_factory=set)


@dataclass
class Node:
    id: str
    cast: dict[str, str] = field(default_factory=dict)
    emails: list[Email] = field(default_factory=list)
    node_depends_on: list[Edge] = field(default_factory=list)


@dataclass
class Corpus:
    nodes: dict[str, Node]
    emails: dict[str, Email]                  # flat, by id
    emission_map: dict[str, str]              # anchor name -> emitting email id

    def topo_order(self) -> list[str]:
        return _topo_sort(self.emails)

    def ancestors(self, email_id: str) -> set[str]:
        seen: set[str] = set()
        stack = [e.email for e in self.emails[email_id].depends_on]
        while stack:
            cur = stack.pop()
            if cur in seen:
                continue
            seen.add(cur)
            stack.extend(e.email for e in self.emails[cur].depends_on)
        return seen


# --- parsing helpers -------------------------------------------------------

def _parse_edge(raw: dict) -> Edge:
    etype = raw.get("type", "static")
    if etype not in EDGE_TYPES:
        raise CorpusError(f"bad edge type {etype!r} (expected one of {sorted(EDGE_TYPES)})")
    if "email" in raw:
        return Edge(email=raw["email"], type=etype)
    if "node" in raw:
        return Edge(email=f"@node:{raw['node']}", type=etype)   # sentinel, expanded later
    raise CorpusError(f"edge missing 'email' or 'node': {raw!r}")


def _parse_expect(raw: dict) -> ExpectEntry:
    action = raw.get("action")
    if action not in VALID_ACTIONS:
        raise CorpusError(f"bad action {action!r} (expected one of {sorted(VALID_ACTIONS)})")
    return ExpectEntry(
        action=action,
        title_match=list(raw.get("title_match", [])),
        start=raw.get("start"),
        due=raw.get("due"),
        count=raw.get("count"),
        tolerance=raw.get("tolerance", "exact_day"),
        recurrence=raw.get("recurrence"),
    )


def _parse_answer(raw: dict) -> Answer:
    return Answer(
        expect=[_parse_expect(e) for e in raw.get("expect", [])],
        forbid=[ForbidEntry(action=f.get("action"), title_match=list(f.get("title_match", [])))
                for f in raw.get("forbid", [])],
        emits=dict(raw.get("emits", {})),
    )


def _refs_in(text: str) -> set[str]:
    return set(_ANCHOR_REF_RE.findall(text))


def _email_predicates(ans: Answer) -> list[str]:
    """Every expression string that appears in an answer's predicates."""
    exprs: list[str] = []
    for e in ans.expect:
        for pred in (e.start, e.due):
            if not pred:
                continue
            for v in pred.values():
                if isinstance(v, list):
                    exprs.extend(str(x) for x in v)
                else:
                    exprs.append(str(v))
        if e.recurrence and "start" in e.recurrence:
            exprs.append(str(e.recurrence["start"]))
    return exprs


def _build_email(node_id: str, raw: dict) -> Email:
    body = raw.get("body", "")
    answer = _parse_answer(raw.get("answer", {}))

    # emissions: body {!name=expr} tokens + explicit answer.emits
    emits: dict[str, str] = {name: expr.strip() for name, expr in _EMIT_IN_BODY_RE.findall(body)}
    emits.update(answer.emits)

    # anchor references: body tokens (minus the emission target) + answer predicates
    refs: set[str] = set()
    for tok in _TOKEN_RE.findall(body):
        refs |= _refs_in(tok)
    for expr in _email_predicates(answer):
        refs |= _refs_in(expr)

    return Email(
        id=raw["id"],
        node=node_id,
        sender=raw.get("from", raw.get("sender", "")),
        recipients=raw.get("to", raw.get("recipients", [])) if isinstance(raw.get("to", raw.get("recipients", [])), list)
                   else [raw.get("to", raw.get("recipients", ""))],
        subject=raw.get("subject", ""),
        body=body,
        depends_on=[_parse_edge(e) for e in raw.get("depends_on", [])],
        answer=answer,
        emits=emits,
        anchor_refs=refs,
    )


# --- loading ---------------------------------------------------------------

def load_corpus(corpus_dir: str | Path) -> Corpus:
    corpus_dir = Path(corpus_dir)
    node_files = sorted((corpus_dir / "nodes").glob("*.json"))
    if not node_files:
        raise CorpusError(f"no node files under {corpus_dir}/nodes/")

    nodes: dict[str, Node] = {}
    emails: dict[str, Email] = {}

    for path in node_files:
        raw = json.loads(path.read_text())
        node_id = raw["id"]
        if node_id in nodes:
            raise CorpusError(f"duplicate node id {node_id!r} ({path})")
        node = Node(
            id=node_id,
            cast=raw.get("cast", {}),
            node_depends_on=[_parse_edge(e) for e in raw.get("node_depends_on", [])],
        )
        for eraw in raw.get("emails", []):
            email = _build_email(node_id, eraw)
            if email.id in emails:
                raise CorpusError(f"duplicate email id {email.id!r} ({path})")
            node.emails.append(email)
            emails[email.id] = email
        nodes[node_id] = node

    _expand_node_edges(nodes, emails)
    emission_map = _build_emission_map(emails)
    corpus = Corpus(nodes=nodes, emails=emails, emission_map=emission_map)
    lint(corpus)
    return corpus


def _expand_node_edges(nodes: dict[str, Node], emails: dict[str, Email]) -> None:
    """node_depends_on sugar -> 'all of ancestor's emails precede all of mine'.

    Also rewrite any per-email '@node:X' sentinel edges the same way.
    """
    by_node: dict[str, list[str]] = {nid: [e.id for e in n.emails] for nid, n in nodes.items()}

    for node in nodes.values():
        for dep in node.node_depends_on:
            anc = dep.email[len("@node:"):] if dep.email.startswith("@node:") else dep.email
            if anc not in by_node:
                raise CorpusError(f"node {node.id!r} depends on unknown node {anc!r}")
            for em in node.emails:
                em.depends_on.extend(Edge(email=aid, type=dep.type) for aid in by_node[anc])

    for email in emails.values():
        expanded: list[Edge] = []
        for dep in email.depends_on:
            if dep.email.startswith("@node:"):
                anc = dep.email[len("@node:"):]
                if anc not in by_node:
                    raise CorpusError(f"email {email.id!r} depends on unknown node {anc!r}")
                expanded.extend(Edge(email=aid, type=dep.type) for aid in by_node[anc])
            else:
                expanded.append(dep)
        email.depends_on = expanded


def _build_emission_map(emails: dict[str, Email]) -> dict[str, str]:
    emission_map: dict[str, str] = {}
    for email in emails.values():
        for name in email.emits:
            if name in emission_map:
                raise CorpusError(
                    f"anchor @{name} emitted by both {emission_map[name]!r} and {email.id!r}"
                )
            emission_map[name] = email.id
    return emission_map


# --- linter ----------------------------------------------------------------

def _topo_sort(emails: dict[str, Email]) -> list[str]:
    indeg = {eid: 0 for eid in emails}
    adj: dict[str, list[str]] = {eid: [] for eid in emails}
    for eid, email in emails.items():
        for dep in email.depends_on:
            if dep.email not in emails:
                raise CorpusError(f"email {eid!r} depends on unknown email {dep.email!r}")
            adj[dep.email].append(eid)
            indeg[eid] += 1
    queue = sorted(eid for eid, d in indeg.items() if d == 0)
    order: list[str] = []
    while queue:
        cur = queue.pop(0)
        order.append(cur)
        for nxt in adj[cur]:
            indeg[nxt] -= 1
            if indeg[nxt] == 0:
                queue.append(nxt)
                queue.sort()
    if len(order) != len(emails):
        cyclic = sorted(set(emails) - set(order))
        raise CorpusError(f"dependency cycle involving: {cyclic}")
    return order


def lint(corpus: Corpus) -> None:
    """Validate the corpus or raise CorpusError. Run automatically by load_corpus."""
    emails = corpus.emails

    # 1. references exist + acyclic (topo_sort checks both)
    _topo_sort(emails)

    # 2. every token / predicate expression parses
    for email in emails.values():
        for tok in _TOKEN_RE.findall(email.body):
            inner = tok.strip()
            m = re.match(r"^!\s*[A-Za-z_]\w*\s*=\s*(.+)$", inner)
            expr = m.group(1) if m else inner
            try:
                resolver._parse_expr(expr.strip())
            except resolver.ResolverError as exc:
                raise CorpusError(f"email {email.id!r}: bad body token {{{tok}}}: {exc}")
        for expr in _email_predicates(email.answer):
            try:
                resolver._parse_expr(expr.strip())
            except resolver.ResolverError as exc:
                raise CorpusError(f"email {email.id!r}: bad answer expression {expr!r}: {exc}")

    # 3. every referenced anchor is emitted by a transitive ancestor
    for email in emails.values():
        if not email.anchor_refs:
            continue
        ancestors = corpus.ancestors(email.id)
        for name in email.anchor_refs:
            src = corpus.emission_map.get(name)
            if src is None:
                raise CorpusError(f"email {email.id!r} references undefined anchor @{name}")
            if src == email.id:
                continue                       # may reference an anchor it emits itself
            if src not in ancestors:
                raise CorpusError(
                    f"email {email.id!r} references @{name} emitted by {src!r}, "
                    f"which is not an ancestor (add a depends_on edge)"
                )

    # 4. an email that references an ancestor's anchor in its ANSWER should carry a
    #    'date' edge to that ancestor (so a serve-by window is derived). Warn-level:
    #    enforce it, since it's the only way the scheduler learns the deadline.
    for email in emails.values():
        answer_refs: set[str] = set()
        for expr in _email_predicates(email.answer):
            answer_refs |= _refs_in(expr)
        for name in answer_refs:
            src = corpus.emission_map.get(name)
            if src is None or src == email.id:
                continue
            has_date_edge = any(
                dep.type == "date" and (dep.email == src or src in corpus.ancestors(dep.email) or dep.email in corpus.ancestors(src))
                for dep in email.depends_on
            ) or any(dep.type == "date" for dep in email.depends_on if dep.email == src)
            if not has_date_edge:
                raise CorpusError(
                    f"email {email.id!r} uses @{name} (from {src!r}) in its answer but has no "
                    f"'date' dependency edge to it — the serve-by window can't be derived"
                )
