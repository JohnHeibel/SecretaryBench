"""Round-trip tests: load the real corpus into forms and recompile losslessly."""
from __future__ import annotations

import json
import shutil

from authoring import corpus_io
from sb import schema

CORPUS = "corpus"


def test_load_graph_is_clean_on_real_corpus():
    graph = corpus_io.load_graph(CORPUS)
    assert graph["errors"] == []
    assert {t["id"] for t in graph["threads"]} == {"acme-client", "henderson", "hr-policy"}
    # acme.request body has one chip that decompiled (not raw)
    req = graph["emails"]["acme-client.request"]
    chip_segs = [s for s in req["body_segments"] if s["type"] == "chip"]
    assert len(chip_segs) == 1
    assert chip_segs[0]["chip"]["base"]["kind"] == "next_weekday"


def test_reachable_anchors_follow_the_dag():
    graph = corpus_io.load_graph(CORPUS)
    # henderson.kickoff depends (transitively) on signing, which emits @signing
    kickoff = graph["emails"]["henderson.kickoff"]
    assert "signing" in kickoff["reachable_anchors"]
    # the intro email sees no anchors yet
    intro = graph["emails"]["henderson.intro"]
    assert intro["reachable_anchors"] == []


def test_edge_badge_inference():
    graph = corpus_io.load_graph(CORPUS)
    kickoff = graph["emails"]["henderson.kickoff"]
    # its answer uses @signing, so its edge back to signing should recommend "date"
    assert any(d["recommended"] == "date" for d in kickoff["depends_on"])


def test_body_segments_recompile_to_identical_string():
    graph = corpus_io.load_graph(CORPUS)
    for raw_path in ["acme-client", "henderson", "hr-policy"]:
        original = json.loads(open(f"corpus/nodes/{raw_path}.json").read())
        for eraw in original.get("emails", []):
            form = graph["emails"][eraw["id"]]
            rebuilt = corpus_io._segments_to_body(form["body_segments"])
            assert rebuilt == eraw.get("body", ""), eraw["id"]


def test_full_thread_recompile_relints_clean(tmp_path):
    # Load, recompile every thread, write to a temp corpus, and confirm the real
    # linter still accepts it (semantic round-trip).
    graph = corpus_io.load_graph(CORPUS)
    dest = tmp_path / "corpus"
    (dest / "nodes").mkdir(parents=True)
    forms_by_thread: dict[str, dict] = {t["id"]: {**t, "emails": []} for t in graph["threads"]}
    for eid, eform in graph["emails"].items():
        forms_by_thread[eform["thread"]]["emails"].append(eform)
    for tid, tform in forms_by_thread.items():
        raw = corpus_io.thread_to_raw(tform)
        (dest / "nodes" / f"{tid}.json").write_text(json.dumps(raw, indent=2))
    # the real loader (with full lint) must accept the recompiled corpus
    loaded = schema.load_corpus(dest)
    assert set(loaded.nodes) == {"acme-client", "henderson", "hr-policy"}
