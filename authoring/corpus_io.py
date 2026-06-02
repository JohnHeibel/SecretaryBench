"""
authoring.corpus_io — tolerant reader/writer between corpus/nodes/*.json and the
editor's structured "form" model.

Two responsibilities:
  * load_graph(dir)  -> a never-fails graph payload for the canvas + editor
                        (threads, email forms with chips, inferred edge badges,
                         per-email pickable anchors, plus any lint errors).
  * write_node(...)  -> compile an edited thread form (chips -> grammar) back
                        into a corpus node JSON file, then lint.

All grammar lives server-side: the UI sends/receives CHIPS, never raw tokens
(except the explicit raw escape-hatch chip), and this module does the
compile/decompile via authoring.chips.
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Optional

from authoring import chips
from sb import schema

_TOKEN_RE = re.compile(r"\{([^{}]*)\}")
_EMIT_RE = re.compile(r"^\s*!\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.+)$")
_FACT_DEF_RE = re.compile(r"^\s*=\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.+?)\s*$")
_FACT_REF_RE = re.compile(r"^\s*=\s*([A-Za-z_][A-Za-z0-9_]*)\s*$")

_EVENT_ACTIONS = {"create_event", "reschedule"}


# ===========================================================================
# READ:  corpus JSON  ->  editor form
# ===========================================================================

def load_graph(corpus_dir: str | Path) -> dict:
    """Build the full editor payload. Never raises on bad corpus content; any
    problems are returned as `errors` so the UI can show a red banner but still
    render what parsed."""
    corpus_dir = Path(corpus_dir)
    node_dir = corpus_dir / "nodes"
    errors: list[str] = []
    raw_nodes: dict[str, dict] = {}

    for path in sorted(node_dir.glob("*.json")):
        try:
            raw = json.loads(path.read_text())
            nid = raw["id"]
            raw_nodes[nid] = raw
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{path.name}: {exc}")

    # email-level edges (expanded from node_depends_on sugar), tolerant
    deps, emits, answer_refs = _build_edges_and_emits(raw_nodes, errors)
    emission_map = _emission_map(emits, errors)
    anc = _ancestors_all(deps)
    fact_emits, fact_map, fact_values = _fact_tables(raw_nodes, errors)

    threads = []
    emails: dict[str, dict] = {}
    for nid, raw in raw_nodes.items():
        email_ids = [e.get("id") for e in raw.get("emails", []) if e.get("id")]
        threads.append({
            "id": nid,
            "cast": raw.get("cast", {}),
            "scenario": raw.get("scenario") or schema.DEFAULT_SCENARIO,
            "node_depends_on": [_edge_dict(e) for e in raw.get("node_depends_on", [])],
            "emails": email_ids,
        })
        for eraw in raw.get("emails", []):
            eid = eraw.get("id")
            if not eid:
                continue
            ancestors = anc.get(eid, set())
            reachable = sorted(n for n, src in emission_map.items() if src in ancestors)
            reach_facts = sorted(n for n, src in fact_map.items()
                                 if src in ancestors or src == eid)
            emails[eid] = _email_to_form(
                nid, eraw,
                ancestors=ancestors, emission_map=emission_map,
                answer_refs=answer_refs.get(eid, set()), reachable=reachable,
                defined_facts=fact_emits.get(eid, {}), reachable_facts=reach_facts,
            )

    return {"threads": threads, "emails": emails,
            "emission_map": emission_map, "fact_map": fact_map, "fact_values": fact_values,
            "errors": errors}


def _email_to_form(node_id: str, eraw: dict, *, ancestors: set[str],
                   emission_map: dict[str, str], answer_refs: set[str],
                   reachable: list[str], defined_facts: dict[str, str],
                   reachable_facts: list[str]) -> dict:
    body = eraw.get("body", "")
    answer = eraw.get("answer", {}) or {}

    deps = [_edge_dict(e) for e in eraw.get("depends_on", [])]
    # recommend a badge: a dep is "deadline" (date) if this email's answer leans
    # on an anchor emitted by that dep (or something upstream of it).
    answer_sources = {emission_map.get(n) for n in answer_refs} - {None}
    for d in deps:
        upstream = {d["email"]} | _reach(d["email"], _DEP_INDEX)
        d["recommended"] = "date" if (answer_sources & upstream) else "static"

    to = eraw.get("to", eraw.get("recipients", []))
    if isinstance(to, str):
        to = [to]

    return {
        "id": eraw["id"],
        "thread": node_id,
        "from": eraw.get("from", eraw.get("sender", "")),
        "to": to,
        "subject": eraw.get("subject", ""),
        "body_segments": _split_body(body),
        "depends_on": deps,
        "answer": _answer_to_form(answer),
        "emits": _body_emits(body),
        "reachable_anchors": reachable,
        "defined_facts": defined_facts,
        "reachable_facts": reachable_facts,
    }


def _split_body(body: str) -> list[dict]:
    """Split a body string into ordered text / chip segments for inline editing."""
    segs: list[dict] = []
    pos = 0
    for m in _TOKEN_RE.finditer(body):
        if m.start() > pos:
            segs.append({"type": "text", "value": body[pos:m.start()]})
        inner = m.group(1).strip()
        fd = _FACT_DEF_RE.match(inner)
        fr = _FACT_REF_RE.match(inner)
        em = _EMIT_RE.match(inner)
        if fd:
            segs.append({"type": "fact", "name": fd.group(1), "value": fd.group(2).strip(), "token": m.group(0)})
        elif fr:
            segs.append({"type": "fact", "name": fr.group(1), "value": None, "token": m.group(0)})
        elif em:
            segs.append({"type": "chip", "chip": chips.parse_token(em.group(2), emit_as=em.group(1)), "token": m.group(0)})
        else:
            segs.append({"type": "chip", "chip": chips.parse_token(inner), "token": m.group(0)})
        pos = m.end()
    if pos < len(body):
        segs.append({"type": "text", "value": body[pos:]})
    return segs


def _answer_to_form(answer: dict) -> dict:
    expect = [_expect_to_form(e) for e in answer.get("expect", [])]
    forbid = [{"action": f.get("action"), "title_match": list(f.get("title_match", []))}
              for f in answer.get("forbid", [])]
    emits = {name: chips.parse_token(expr) for name, expr in (answer.get("emits", {}) or {}).items()}
    facts = {name: str(v) for name, v in (answer.get("facts", {}) or {}).items()}
    return {"expect": expect, "forbid": forbid, "emits": emits, "facts": facts}


def _expect_to_form(e: dict) -> dict:
    action = e.get("action", "create_event")
    pred = e.get("start") if action in _EVENT_ACTIONS else (e.get("due") or e.get("start"))
    return {
        "action": action,
        "title_match": list(e.get("title_match", [])),
        "when": _predicate_to_form(pred),
        "duration": e.get("duration"),
        "count": e.get("count"),
        "tolerance": e.get("tolerance", "exact_day"),
    }


def _predicate_to_form(pred: Optional[dict]) -> Optional[dict]:
    if not pred:
        return None
    if "eq" in pred:
        return {"match": "on", "chip": chips.parse_token(str(pred["eq"]))}
    if "by" in pred:
        return {"match": "by", "chip": chips.parse_token(str(pred["by"]))}
    if "any_of" in pred:
        return {"match": "any_of", "chips": [chips.parse_token(str(x)) for x in pred["any_of"]]}
    if "in" in pred:
        out = {"match": "within", "chip": chips.parse_token(str(pred["in"]))}
        if "not_in" in pred:
            out["avoid_chip"] = chips.parse_token(str(pred["not_in"]))
        return out
    return None


def _body_emits(body: str) -> dict[str, dict]:
    out: dict[str, dict] = {}
    for m in _TOKEN_RE.finditer(body):
        em = _EMIT_RE.match(m.group(1).strip())
        if em:
            out[em.group(1)] = chips.parse_token(em.group(2), emit_as=em.group(1))
    return out


# ===========================================================================
# WRITE:  editor form  ->  corpus JSON
# ===========================================================================

def thread_to_raw(form: dict) -> dict:
    """Compile a thread form (with chips) into a corpus node dict."""
    raw: dict[str, Any] = {"id": form["id"]}
    if form.get("cast"):
        raw["cast"] = form["cast"]
    # Persist scenario only when it's a real grouping — omit the default bucket so
    # untouched corpora keep their original, scenario-less node files.
    scenario = form.get("scenario")
    if scenario and scenario != schema.DEFAULT_SCENARIO:
        raw["scenario"] = scenario
    if form.get("node_depends_on"):
        raw["node_depends_on"] = [_edge_out(e) for e in form["node_depends_on"]]
    raw["emails"] = [_email_to_raw(e) for e in form.get("emails", [])]
    return raw


def _email_to_raw(form: dict) -> dict:
    out: dict[str, Any] = {
        "id": form["id"],
        "from": form.get("from", ""),
        "to": form.get("to", []),
        "subject": form.get("subject", ""),
        "body": _segments_to_body(form.get("body_segments", [])),
    }
    if form.get("depends_on"):
        out["depends_on"] = [_edge_out(e) for e in form["depends_on"]]
    out["answer"] = _answer_to_raw(form.get("answer", {}))
    return out


def _segments_to_body(segments: list[dict]) -> str:
    parts: list[str] = []
    for seg in segments:
        kind = seg.get("type")
        if kind == "text":
            parts.append(seg.get("value", ""))
        elif kind == "chip":
            parts.append(chips.compile_body_token(seg["chip"]))
        elif kind == "fact":
            if seg.get("value") is None:
                parts.append(f"{{={seg['name']}}}")
            else:
                parts.append(f"{{={seg['name']} = {seg['value']}}}")
    return "".join(parts)


def _answer_to_raw(answer: dict) -> dict:
    out: dict[str, Any] = {}
    expect = [_expect_to_raw(e) for e in answer.get("expect", [])]
    if expect:
        out["expect"] = expect
    forbid = [{k: v for k, v in (("action", f.get("action")),
                                 ("title_match", f.get("title_match") or [])) if v}
              for f in answer.get("forbid", [])]
    if forbid:
        out["forbid"] = forbid
    emits = {name: chips.compile_chip(chip) for name, chip in (answer.get("emits", {}) or {}).items()}
    if emits:
        out["emits"] = emits
    facts = {name: str(v) for name, v in (answer.get("facts", {}) or {}).items()}
    if facts:
        out["facts"] = facts
    return out


def _expect_to_raw(form: dict) -> dict:
    action = form.get("action", "create_event")
    out: dict[str, Any] = {"action": action}
    if form.get("title_match"):
        out["title_match"] = form["title_match"]
    pred = _predicate_to_raw(form.get("when"))
    if pred is not None:
        out["start" if action in _EVENT_ACTIONS else "due"] = pred
    if form.get("duration"):
        out["duration"] = form["duration"]
    if form.get("count") is not None:
        out["count"] = form["count"]
    if form.get("tolerance"):
        out["tolerance"] = form["tolerance"]
    return out


def _predicate_to_raw(pred: Optional[dict]) -> Optional[dict]:
    if not pred:
        return None
    match = pred.get("match")
    if match == "on":
        return {"eq": chips.compile_chip(pred["chip"])}
    if match == "by":
        return {"by": chips.compile_chip(pred["chip"])}
    if match == "any_of":
        return {"any_of": [chips.compile_chip(c) for c in pred.get("chips", [])]}
    if match == "within":
        out = {"in": chips.compile_chip(pred["chip"])}
        if pred.get("avoid_chip"):
            out["not_in"] = chips.compile_chip(pred["avoid_chip"])
        return out
    return None


def write_node(corpus_dir: str | Path, thread_form: dict) -> list[str]:
    """Compile and write one thread to corpus/nodes/<id>.json, then lint the
    whole corpus. Returns lint errors (empty list == clean)."""
    corpus_dir = Path(corpus_dir)
    raw = thread_to_raw(thread_form)
    path = corpus_dir / "nodes" / f"{raw['id']}.json"
    path.write_text(json.dumps(raw, indent=2) + "\n")
    return validate(corpus_dir)


def delete_node(corpus_dir: str | Path, thread_id: str) -> list[str]:
    path = Path(corpus_dir) / "nodes" / f"{thread_id}.json"
    if path.exists():
        path.unlink()
    return validate(corpus_dir)


def validate(corpus_dir: str | Path) -> list[str]:
    """Run the real schema linter; return error strings (empty == clean)."""
    try:
        schema.load_corpus(corpus_dir)
        return []
    except schema.CorpusError as exc:
        return [str(exc)]
    except Exception as exc:  # noqa: BLE001
        return [f"{type(exc).__name__}: {exc}"]


# ===========================================================================
# tolerant graph helpers (no linting, never raise)
# ===========================================================================

_DEP_INDEX: dict[str, list[str]] = {}   # module-level cache used by _reach during a load


def _build_edges_and_emits(raw_nodes: dict[str, dict], errors: list[str]):
    """Return (deps: email->[dep_email], emits: email->{name:expr}, answer_refs)."""
    by_node = {nid: [e.get("id") for e in raw.get("emails", []) if e.get("id")]
               for nid, raw in raw_nodes.items()}
    deps: dict[str, list[str]] = {}
    emits: dict[str, dict[str, str]] = {}
    answer_refs: dict[str, set[str]] = {}

    for nid, raw in raw_nodes.items():
        node_deps = raw.get("node_depends_on", [])
        for eraw in raw.get("emails", []):
            eid = eraw.get("id")
            if not eid:
                continue
            edge_list: list[str] = []
            # node-level sugar -> all emails of each ancestor node
            for nd in node_deps:
                anc = nd.get("node")
                edge_list += [x for x in by_node.get(anc, []) if x]
            for d in eraw.get("depends_on", []):
                if "email" in d:
                    edge_list.append(d["email"])
                elif "node" in d:
                    edge_list += [x for x in by_node.get(d["node"], []) if x]
            deps[eid] = edge_list

            # emits from body + explicit answer.emits
            em: dict[str, str] = {}
            for m in _TOKEN_RE.finditer(eraw.get("body", "")):
                me = _EMIT_RE.match(m.group(1).strip())
                if me:
                    em[me.group(1)] = me.group(2).strip()
            for name, expr in (eraw.get("answer", {}).get("emits", {}) or {}).items():
                em[name] = str(expr)
            emits[eid] = em

            # anchor refs used in the ANSWER (drives edge badge inference)
            refs: set[str] = set()
            for expr in _answer_predicate_exprs(eraw.get("answer", {})):
                refs |= set(re.findall(r"@([A-Za-z_]\w*)", expr))
            answer_refs[eid] = refs

    global _DEP_INDEX
    _DEP_INDEX = deps
    return deps, emits, answer_refs


def _answer_predicate_exprs(answer: dict) -> list[str]:
    out: list[str] = []
    for e in answer.get("expect", []):
        for key in ("start", "due"):
            pred = e.get(key)
            if not pred:
                continue
            for v in pred.values():
                if isinstance(v, list):
                    out += [str(x) for x in v]
                else:
                    out.append(str(v))
    return out


def _fact_tables(raw_nodes: dict[str, dict], errors: list[str]):
    """Return (per-email defined facts, fact name -> email, fact name -> value)."""
    per_email: dict[str, dict[str, str]] = {}
    fact_map: dict[str, str] = {}
    values: dict[str, str] = {}
    for raw in raw_nodes.values():
        for eraw in raw.get("emails", []):
            eid = eraw.get("id")
            if not eid:
                continue
            defined: dict[str, str] = {}
            for m in _TOKEN_RE.finditer(eraw.get("body", "")):
                fd = _FACT_DEF_RE.match(m.group(1).strip())
                if fd:
                    defined[fd.group(1)] = fd.group(2).strip()
            for name, val in (eraw.get("answer", {}).get("facts", {}) or {}).items():
                defined[name] = str(val)
            per_email[eid] = defined
            for name, val in defined.items():
                if name in fact_map:
                    errors.append(f"fact ={name} defined by both {fact_map[name]!r} and {eid!r}")
                    continue
                fact_map[name] = eid
                values[name] = val
    return per_email, fact_map, values


def _emission_map(emits: dict[str, dict[str, str]], errors: list[str]) -> dict[str, str]:
    out: dict[str, str] = {}
    for eid, names in emits.items():
        for name in names:
            if name in out:
                errors.append(f"anchor @{name} emitted by both {out[name]!r} and {eid!r}")
                continue
            out[name] = eid
    return out


def _ancestors_all(deps: dict[str, list[str]]) -> dict[str, set[str]]:
    out: dict[str, set[str]] = {}
    for eid in deps:
        out[eid] = _reach(eid, deps)
    return out


def _reach(eid: str, deps: dict[str, list[str]]) -> set[str]:
    seen: set[str] = set()
    stack = list(deps.get(eid, []))
    while stack:
        cur = stack.pop()
        if cur in seen:
            continue
        seen.add(cur)
        stack.extend(deps.get(cur, []))
    return seen


def _edge_dict(e: dict) -> dict:
    out = {"type": e.get("type", "static")}
    if "email" in e:
        out["email"] = e["email"]
    elif "node" in e:
        out["node"] = e["node"]
    return out


def _edge_out(e: dict) -> dict:
    out: dict[str, Any] = {}
    if e.get("email"):
        out["email"] = e["email"]
    elif e.get("node"):
        out["node"] = e["node"]
    out["type"] = e.get("type", "static")
    return out
