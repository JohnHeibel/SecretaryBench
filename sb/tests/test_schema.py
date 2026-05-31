"""Tests for the corpus loader, node-edge expansion, and linter."""
import json

import pytest

from sb.schema import CorpusError, load_corpus

CORPUS = "corpus"


def test_pilot_corpus_loads_and_lints():
    c = load_corpus(CORPUS)
    assert set(c.nodes) == {"henderson", "hr-policy", "acme-client"}
    assert len(c.emails) == 6


def test_emission_map_built():
    c = load_corpus(CORPUS)
    assert c.emission_map["signing"] == "henderson.signing"


def test_node_sugar_expands_to_email_edges():
    c = load_corpus(CORPUS)
    # acme-client node_depends_on hr-policy -> every acme email depends on the memo.
    req = c.emails["acme-client.request"]
    assert any(d.email == "hr-policy.memo" and d.type == "static" for d in req.depends_on)


def test_ancestors_transitive():
    c = load_corpus(CORPUS)
    anc = c.ancestors("henderson.kickoff")
    assert anc == {"henderson.signing", "henderson.intro"}


def test_topo_order_respects_dependencies():
    c = load_corpus(CORPUS)
    order = c.topo_order()
    assert order.index("henderson.intro") < order.index("henderson.signing")
    assert order.index("henderson.signing") < order.index("henderson.kickoff")


# --- linter rejects bad corpora --------------------------------------------

def _write_node(tmp_path, name, obj):
    (tmp_path / "nodes").mkdir(exist_ok=True)
    (tmp_path / "nodes" / f"{name}.json").write_text(json.dumps(obj))


def test_cycle_rejected(tmp_path):
    _write_node(tmp_path, "n", {
        "id": "n",
        "emails": [
            {"id": "n.a", "depends_on": [{"email": "n.b", "type": "static"}], "answer": {"expect": []}},
            {"id": "n.b", "depends_on": [{"email": "n.a", "type": "static"}], "answer": {"expect": []}},
        ],
    })
    with pytest.raises(CorpusError, match="cycle"):
        load_corpus(tmp_path)


def test_undefined_anchor_rejected(tmp_path):
    _write_node(tmp_path, "n", {
        "id": "n",
        "emails": [
            {"id": "n.a", "body": "meet on {@ghost+1d}", "depends_on": [],
             "answer": {"expect": []}},
        ],
    })
    with pytest.raises(CorpusError, match="undefined anchor"):
        load_corpus(tmp_path)


def test_anchor_without_date_edge_rejected(tmp_path):
    # n.b uses @x in its ANSWER but only has a static edge -> no derivable window.
    _write_node(tmp_path, "n", {
        "id": "n",
        "emails": [
            {"id": "n.a", "body": "set {!x = serve+3d}", "depends_on": [], "answer": {"expect": []}},
            {"id": "n.b", "depends_on": [{"email": "n.a", "type": "static"}],
             "answer": {"expect": [
                 {"action": "create_event", "title_match": ["k"], "start": {"eq": "@x+1w"}}]}},
        ],
    })
    with pytest.raises(CorpusError, match="serve-by window"):
        load_corpus(tmp_path)


def test_anchor_in_answer_with_date_edge_ok(tmp_path):
    _write_node(tmp_path, "n", {
        "id": "n",
        "emails": [
            {"id": "n.a", "body": "set {!x = serve+3d}", "depends_on": [], "answer": {"expect": []}},
            {"id": "n.b", "depends_on": [{"email": "n.a", "type": "date"}],
             "answer": {"expect": [
                 {"action": "create_event", "title_match": ["k"], "start": {"eq": "@x+1w"}}]}},
        ],
    })
    c = load_corpus(tmp_path)
    assert c.emission_map["x"] == "n.a"
