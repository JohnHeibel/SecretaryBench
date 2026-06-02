"""Tests for the corpus loader, node-edge expansion, and linter."""
import json
from pathlib import Path

import pytest

from sb.schema import CorpusError, build_corpus, load_corpus

FIXTURE = str(Path(__file__).parent / "fixtures")


def test_fixture_corpus_loads_and_lints():
    c = load_corpus(FIXTURE)
    assert set(c.nodes) == {"alpha", "beta", "gamma"}
    assert len(c.emails) == 6


def test_emission_map_built():
    c = load_corpus(FIXTURE)
    assert c.emission_map["brief"] == "alpha.brief"


def test_node_sugar_expands_to_email_edges():
    c = load_corpus(FIXTURE)
    # beta node_depends_on gamma -> every beta email depends on the gamma notice.
    book = c.emails["beta.book"]
    assert any(d.email == "gamma.notice" and d.type == "static" for d in book.depends_on)


def test_ancestors_transitive():
    c = load_corpus(FIXTURE)
    anc = c.ancestors("alpha.review")
    assert anc == {"alpha.brief", "alpha.intro"}


def test_topo_order_respects_dependencies():
    c = load_corpus(FIXTURE)
    order = c.topo_order()
    assert order.index("alpha.intro") < order.index("alpha.brief")
    assert order.index("alpha.brief") < order.index("alpha.review")


# --- linter rejects bad corpora --------------------------------------------

def _write_node(tmp_path, name, obj):
    (tmp_path / "nodes").mkdir(exist_ok=True)
    (tmp_path / "nodes" / f"{name}.json").write_text(json.dumps(obj))


def test_duplicate_node_id_rejected():
    # two node files declaring the same `id` must be refused — names are identities.
    node = {"id": "dup", "emails": [{"id": "dup.a", "answer": {"ops": []}}]}
    with pytest.raises(CorpusError, match="duplicate node id"):
        build_corpus([node, {"id": "dup", "emails": [{"id": "dup.b", "answer": {"ops": []}}]}])


def test_colliding_match_keywords_rejected(tmp_path):
    # two event obligations in one node whose keyword filters catch each other.
    _write_node(tmp_path, "n", {
        "id": "n",
        "emails": [
            {"id": "n.a", "answer": {"ops": [
                {"create": "kickoff", "kind": "event", "on": {"eq": "serve+1w"}, "match": ["meeting"]}]}},
            {"id": "n.b", "answer": {"ops": [
                {"create": "review", "kind": "event", "on": {"eq": "serve+2w"}, "match": ["meeting"]}]}},
        ],
    })
    with pytest.raises(CorpusError, match="colliding match keywords"):
        load_corpus(tmp_path)


def test_substring_match_keywords_rejected(tmp_path):
    # "kickoff" is a substring of "kickoff sync", so the broader filter is caught too.
    _write_node(tmp_path, "n", {
        "id": "n",
        "emails": [
            {"id": "n.a", "answer": {"ops": [
                {"create": "k1", "kind": "event", "on": {"eq": "serve+1w"}, "match": ["kickoff"]}]}},
            {"id": "n.b", "answer": {"ops": [
                {"create": "k2", "kind": "event", "on": {"eq": "serve+2w"}, "match": ["kickoff", "sync"]}]}},
        ],
    })
    with pytest.raises(CorpusError, match="colliding match keywords"):
        load_corpus(tmp_path)


def test_overlapping_but_distinct_keywords_ok(tmp_path):
    # shared "meeting" but each has a distinguishing word -> neither filter catches the
    # other -> must NOT be flagged (no false positive on safe overlap).
    _write_node(tmp_path, "n", {
        "id": "n",
        "emails": [
            {"id": "n.a", "answer": {"ops": [
                {"create": "board", "kind": "event", "on": {"eq": "serve+1w"}, "match": ["board", "meeting"]}]}},
            {"id": "n.b", "answer": {"ops": [
                {"create": "staff", "kind": "event", "on": {"eq": "serve+2w"}, "match": ["staff", "meeting"]}]}},
        ],
    })
    load_corpus(tmp_path)   # should not raise


def test_same_keyword_different_kind_ok(tmp_path):
    # an event and a todo never share a pool, so identical keywords don't collide.
    _write_node(tmp_path, "n", {
        "id": "n",
        "emails": [
            {"id": "n.a", "answer": {"ops": [
                {"create": "ev", "kind": "event", "on": {"eq": "serve+1w"}, "match": ["sync"]}]}},
            {"id": "n.b", "answer": {"ops": [
                {"create": "td", "kind": "todo", "on": {"by": "serve+2w"}, "match": ["sync"]}]}},
        ],
    })
    load_corpus(tmp_path)   # should not raise


def test_cycle_rejected(tmp_path):
    _write_node(tmp_path, "n", {
        "id": "n",
        "emails": [
            {"id": "n.a", "depends_on": [{"email": "n.b", "type": "static"}], "answer": {"ops": []}},
            {"id": "n.b", "depends_on": [{"email": "n.a", "type": "static"}], "answer": {"ops": []}},
        ],
    })
    with pytest.raises(CorpusError, match="cycle"):
        load_corpus(tmp_path)


def test_undefined_anchor_rejected(tmp_path):
    _write_node(tmp_path, "n", {
        "id": "n",
        "emails": [
            {"id": "n.a", "body": "meet on {@ghost+1d}", "depends_on": [],
             "answer": {"ops": []}},
        ],
    })
    with pytest.raises(CorpusError, match="undefined anchor"):
        load_corpus(tmp_path)


def test_anchor_without_date_edge_rejected(tmp_path):
    # n.b uses @x in its ANSWER but only has a static edge -> no derivable window.
    _write_node(tmp_path, "n", {
        "id": "n",
        "emails": [
            {"id": "n.a", "body": "set {!x = serve+3d}", "depends_on": [], "answer": {"ops": []}},
            {"id": "n.b", "depends_on": [{"email": "n.a", "type": "static"}],
             "answer": {"ops": [
                 {"create": "k", "kind": "event", "on": {"eq": "@x+1w"}}]}},
        ],
    })
    with pytest.raises(CorpusError, match="serve-by window"):
        load_corpus(tmp_path)


def test_anchor_in_answer_with_date_edge_ok(tmp_path):
    _write_node(tmp_path, "n", {
        "id": "n",
        "emails": [
            {"id": "n.a", "body": "set {!x = serve+3d}", "depends_on": [], "answer": {"ops": []}},
            {"id": "n.b", "depends_on": [{"email": "n.a", "type": "date"}],
             "answer": {"ops": [
                 {"create": "k", "kind": "event", "on": {"eq": "@x+1w"}}]}},
        ],
    })
    c = load_corpus(tmp_path)
    assert c.emission_map["x"] == "n.a"


def test_move_derives_edge_to_creator(tmp_path):
    # a move references its obligation by name; the dependency edge auto-derives.
    _write_node(tmp_path, "n", {
        "id": "n",
        "emails": [
            {"id": "n.make", "depends_on": [],
             "answer": {"ops": [{"create": "call", "kind": "event", "on": {"eq": "next:THU"}}]}},
            {"id": "n.move", "depends_on": [],
             "answer": {"ops": [{"move": "call", "on": {"eq": "next:MON"}}]}},
        ],
    })
    c = load_corpus(tmp_path)
    # no explicit depends_on, yet n.move must follow n.make
    assert any(d.email == "n.make" for d in c.emails["n.move"].depends_on)
    order = c.topo_order()
    assert order.index("n.make") < order.index("n.move")
