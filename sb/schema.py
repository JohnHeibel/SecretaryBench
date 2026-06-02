"""
sb.schema — the corpus data model, JSON loader, and linter.

A corpus is a set of node files (corpus/nodes/*.json). A node groups emails that
share a cast and emitted anchors, but all scheduling is computed from a single
FLAT email-level DAG (see SERVING_AND_SCHEMA.md §0). This module loads the files,
flattens node-level sugar into email edges, wires up obligations, derives the
anchor emission map, and lints the result (acyclic, anchors reachable, tokens parse).

Answer keys are VERB-based (ADR 0001): an email's answer is a list of `ops`, each a
`create` / `move` / `cancel` on a NAMED obligation. The name is the obligation's
identity AND a node-scoped anchor — so a later `move`/`cancel` references it by name,
its dependency edge derives, and its date is reusable as `@name`. No `count` field.
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

VERBS = {"create", "move", "cancel"}
OBJ_KINDS = {"event", "todo"}
EDGE_TYPES = {"static", "date"}
_PREDICATE_KEYS = {"eq", "by", "in", "not_in", "any_of", "start"}


class CorpusError(ValueError):
    """Raised when a corpus fails to load or lint."""


@dataclass
class Edge:
    email: str            # prerequisite email id
    type: str             # "static" | "date"


@dataclass
class Op:
    """One verb on a named obligation. The name is the obligation's identity and a
    node-scoped anchor; `match` is the fuzzy title keyword set the grader uses to
    tie the model's calendar object to this obligation (defaults to [name])."""
    verb: str                                   # "create" | "move" | "cancel"
    name: str                                   # obligation identity (also a @name anchor)
    kind: Optional[str] = None                  # "event" | "todo" (set on create; filled in for move/cancel)
    match: list[str] = field(default_factory=list)   # title keywords; defaults to [name]
    on: Optional[dict] = None                   # date predicate {"eq"|"by"|"in"|"any_of"|"not_in": expr}
    tolerance: str = "exact_day"                # exact_day | within:Nd  (whole-day granularity)


@dataclass
class Answer:
    ops: list[Op] = field(default_factory=list)
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
    emits: dict[str, str] = field(default_factory=dict)    # name -> expr (body + answer.emits + obligation anchors)
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


def _parse_op(raw: dict) -> Op:
    verbs_present = [v for v in ("create", "move", "cancel") if v in raw]
    if len(verbs_present) != 1:
        raise CorpusError(f"op must have exactly one of create/move/cancel: {raw!r}")
    verb = verbs_present[0]
    name = raw[verb]
    if not isinstance(name, str) or not name:
        raise CorpusError(f"op {verb} needs a non-empty obligation name: {raw!r}")
    kind = raw.get("kind")
    on = raw.get("on")
    if verb == "create":
        if kind not in OBJ_KINDS:
            raise CorpusError(f"create op {name!r} needs kind {sorted(OBJ_KINDS)}, got {kind!r}")
        if not on:
            raise CorpusError(f"create op {name!r} needs an 'on' date predicate")
    if verb == "move" and not on:
        raise CorpusError(f"move op {name!r} needs an 'on' date predicate")
    if verb == "cancel" and (on or kind):
        raise CorpusError(f"cancel op {name!r} takes neither 'on' nor 'kind'")
    return Op(verb=verb, name=name, kind=kind,
              match=list(raw.get("match", [])) or [name],
              on=on, tolerance=raw.get("tolerance", "exact_day"))


def _parse_answer(raw: dict) -> Answer:
    return Answer(
        ops=[_parse_op(o) for o in raw.get("ops", [])],
        emits=dict(raw.get("emits", {})),
    )


def _refs_in(text: str) -> set[str]:
    return set(_ANCHOR_REF_RE.findall(str(text)))


def _predicate_exprs(predicate: Optional[dict]) -> list[str]:
    """Every expression string in a single date predicate."""
    if not predicate:
        return []
    out: list[str] = []
    for v in predicate.values():
        if isinstance(v, list):
            out.extend(str(x) for x in v)
        else:
            out.append(str(v))
    return out


def _email_predicates(ans: Answer) -> list[str]:
    """Every date expression that appears in an answer's ops (used by the scheduler)."""
    exprs: list[str] = []
    for op in ans.ops:
        exprs.extend(_predicate_exprs(op.on))
    return exprs


def _build_email(node_id: str, raw: dict) -> Email:
    body = raw.get("body", "")
    answer = _parse_answer(raw.get("answer", {}))

    # emissions: body {!name=expr} tokens + explicit answer.emits (obligation anchors added in _wire_obligations)
    emits: dict[str, str] = {name: expr.strip() for name, expr in _EMIT_IN_BODY_RE.findall(body)}
    emits.update(answer.emits)

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
        anchor_refs=set(),     # computed in _recompute_refs after obligation wiring
    )


# --- loading ---------------------------------------------------------------

def load_corpus(corpus_dir: str | Path) -> Corpus:
    corpus_dir = Path(corpus_dir)
    node_files = sorted((corpus_dir / "nodes").glob("*.json"))
    if not node_files:
        raise CorpusError(f"no node files under {corpus_dir}/nodes/")
    raw_nodes = [json.loads(path.read_text()) for path in node_files]
    sources = [str(path) for path in node_files]
    return build_corpus(raw_nodes, sources=sources)


# ids become filenames (corpus/nodes/<id>.json via sb.sync) and are interpolated
# into paths, so they must be a safe, simple slug — no path separators, no `..`.
_ID_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]*$")


def _check_id(value: str, kind: str, src: str) -> None:
    if not isinstance(value, str) or not _ID_RE.match(value) or ".." in value:
        raise CorpusError(f"invalid {kind} id {value!r} ({src}): must match {_ID_RE.pattern} and contain no '..'")


def build_corpus(raw_nodes: list[dict], sources: Optional[list[str]] = None) -> Corpus:
    """Build + lint a Corpus from in-memory raw node dicts (same shape as the node
    JSON files). `sources` is an optional parallel list of human labels (e.g. file
    paths) used only in duplicate-id error messages. This is the disk-free entry
    point the web app's lint endpoint calls; load_corpus reads files then delegates."""
    if sources is None:
        sources = [str(raw.get("id", f"#{i}")) for i, raw in enumerate(raw_nodes)]

    nodes: dict[str, Node] = {}
    emails: dict[str, Email] = {}

    for raw, src in zip(raw_nodes, sources):
        node_id = raw["id"]
        _check_id(node_id, "node", src)
        if node_id in nodes:
            raise CorpusError(f"duplicate node id {node_id!r} ({src})")
        node = Node(
            id=node_id,
            cast=raw.get("cast", {}),
            node_depends_on=[_parse_edge(e) for e in raw.get("node_depends_on", [])],
        )
        for eraw in raw.get("emails", []):
            email = _build_email(node_id, eraw)
            _check_id(email.id, "email", src)
            if email.id in emails:
                raise CorpusError(f"duplicate email id {email.id!r} ({src})")
            node.emails.append(email)
            emails[email.id] = email
        nodes[node_id] = node

    _expand_node_edges(nodes, emails)
    body_anchors = {name for em in emails.values() for name in em.emits}   # before obligation wiring
    _wire_obligations(nodes, emails, body_anchors)
    _recompute_refs(emails)
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


# --- obligations -----------------------------------------------------------

def _obl_anchor(node_id: str, name: str) -> str:
    """The internal, globally-unique anchor name an obligation publishes. Authors
    write the bare name (@kickoff); this node-qualifies it so two nodes can both
    have a 'kickoff' without colliding in the global anchor table."""
    slug = re.sub(r"[^A-Za-z0-9_]", "_", node_id)
    return f"__obl_{slug}__{name}"


def _rewrite_expr(expr: object, rename: dict[str, str]) -> object:
    if not isinstance(expr, str):
        return expr
    return _ANCHOR_REF_RE.sub(lambda m: f"@{rename.get(m.group(1), m.group(1))}", expr)


def _rewrite_predicate(predicate: dict, rename: dict[str, str]) -> dict:
    out: dict = {}
    for k, v in predicate.items():
        if isinstance(v, list):
            out[k] = [_rewrite_expr(x, rename) for x in v]
        elif k in _PREDICATE_KEYS:
            out[k] = _rewrite_expr(v, rename)
        else:
            out[k] = v
    return out


def _predicate_target_expr(predicate: dict) -> Optional[object]:
    """The single expression that anchors an obligation's date (for @name emission)."""
    for key in ("eq", "by", "start", "in"):
        if key in predicate:
            return predicate[key]
    if "any_of" in predicate and predicate["any_of"]:
        return predicate["any_of"][0]
    return None


def _wire_obligations(nodes: dict[str, Node], emails: dict[str, Email], body_anchors: set[str]) -> None:
    """Resolve the verb model into anchors + edges the scheduler/grader already speak:

      - each `create: X` registers obligation X (node-scoped) and, if a sibling op
        references @X, publishes @X as a node-qualified anchor at the creator's serve;
      - `move`/`cancel: X` inherit X's kind/match, and derive a depends_on edge to X's
        creator (a `date` edge when the op references an anchor, else `static`);
      - bare @X references to a sibling obligation are rewritten to the qualified name.
    """
    for node in nodes.values():
        registry: dict[str, dict] = {}   # name -> {creator, kind, match, on}
        for em in node.emails:
            for op in em.answer.ops:
                if op.verb == "create":
                    if op.name in registry:
                        raise CorpusError(f"node {node.id!r}: obligation {op.name!r} created twice")
                    registry[op.name] = {"creator": em.id, "kind": op.kind, "match": op.match, "on": op.on}

        # move/cancel must target an existing obligation; inherit its kind + match
        for em in node.emails:
            for op in em.answer.ops:
                if op.verb in ("move", "cancel"):
                    info = registry.get(op.name)
                    if info is None:
                        raise CorpusError(
                            f"email {em.id!r}: {op.verb} of unknown obligation {op.name!r} "
                            f"(no create in node {node.id!r})")
                    op.kind = info["kind"]
                    if op.match == [op.name]:
                        op.match = info["match"]

        # which sibling obligations are referenced as @name (and aren't a body anchor)?
        siblings = set(registry)
        referenced: set[str] = set()
        for em in node.emails:
            for op in em.answer.ops:
                for expr in _predicate_exprs(op.on):
                    referenced |= {r for r in _refs_in(expr) if r in siblings and r not in body_anchors}

        rename = {name: _obl_anchor(node.id, name) for name in referenced}

        # rewrite op predicates so @name resolves to the obligation's published anchor
        if rename:
            for em in node.emails:
                for op in em.answer.ops:
                    if op.on:
                        op.on = _rewrite_predicate(op.on, rename)

        # the creator publishes @name = its obligation's date (only when referenced)
        for name in referenced:
            info = registry[name]
            target = _predicate_target_expr(info["on"])
            if target is not None:
                creator = emails[info["creator"]]
                creator.emits[rename[name]] = str(_rewrite_expr(target, rename)).strip()

        # derive the dependency edge from each move/cancel to its obligation's creator
        for em in node.emails:
            for op in em.answer.ops:
                if op.verb not in ("move", "cancel"):
                    continue
                creator_id = registry[op.name]["creator"]
                if creator_id == em.id or any(d.email == creator_id for d in em.depends_on):
                    continue
                has_anchor = any(_refs_in(expr) for expr in _predicate_exprs(op.on))
                em.depends_on.append(Edge(email=creator_id, type="date" if has_anchor else "static"))


def _recompute_refs(emails: dict[str, Email]) -> None:
    """Anchor references = body tokens + (post-rewrite) answer predicate expressions."""
    for email in emails.values():
        refs: set[str] = set()
        for tok in _TOKEN_RE.findall(email.body):
            refs |= _refs_in(tok)
        for expr in _email_predicates(email.answer):
            refs |= _refs_in(expr)
        email.anchor_refs = refs


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
    #    'date' edge to that ancestor (so a serve-by window is derived). Move/cancel
    #    edges are auto-derived; this still catches a create that forgot its date edge.
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
