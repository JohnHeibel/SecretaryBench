"""
C-11: `sb.regrade` must not grade a capture against a corpus that changes what the
model was served or when it was served it.

Every changed corpus is a copy in tmp_path of a capture's frozen corpus, so these
tests do not move when `corpus/` is repaired. Nothing under `corpus/` or `captures/`
is ever written; capture-side cases copy the capture first.
"""
import json
import shutil
from pathlib import Path

import pytest

from sb.live.runner import _corpus_hash, _served_email
from sb.regrade import (ReferenceMismatch, _plan, check_reference, load_capture, main,
                        regrade, score)
from sb.schema import load_corpus

C1C = "captures/baseline-sonnet-4-5"
R2 = "captures/baseline-sonnet-4-5-r2"
CAPTURES = [C1C, R2]
SCORES = {C1C: (114, 167), R2: (123, 167)}
GET_EMAIL_FLOOR = {C1C: 148, R2: 167}      # get_email calls in the raw streams (1c: by day 50)

# A non-date-edge create, passed in both runs, whose name nothing else references.
KEY_EMAIL = "Planning.memo-before-the-board-sits"
# A date-edge email whose deadline feeds the scheduler (sb/scheduler.py _anchor_exprs).
PLAN_EMAIL = "Partnership-with-deeptech-companies.new-meeting-during-boston-trip"
BODY_EMAIL = "Company-Retreat.company-retreat-date-conformed"
LEAF_EMAIL = "Innovation-comp.quick-win-to-share"           # no action, no tokens, no dependents


def _frozen_copy(tmp_path, capture=C1C) -> Path:
    dst = tmp_path / "given"
    shutil.copytree(Path(capture) / "corpus", dst)
    return dst


def _node_path(corpus_dir: Path, eid: str) -> Path:
    return corpus_dir / "nodes" / f"{eid.split('.')[0]}.json"


def _edit(corpus_dir: Path, eid: str, change) -> None:
    """Apply `change(email_dict, node_dict)` to one email, keeping the file's format."""
    path = _node_path(corpus_dir, eid)
    node = json.loads(path.read_text())
    change(next(e for e in node["emails"] if e["id"] == eid), node)
    path.write_text(json.dumps(node, indent=2) + "\n")


def _capture_copy(tmp_path, capture=C1C, snapshot=True) -> Path:
    dst = tmp_path / "capture"
    dst.mkdir()
    shutil.copy(Path(capture) / "manifest.json", dst / "manifest.json")
    shutil.copytree(Path(capture) / "days", dst / "days")
    if snapshot:
        shutil.copytree(Path(capture) / "corpus", dst / "corpus")
    return dst


def _reformatted(tmp_path) -> Path:
    given = _frozen_copy(tmp_path)
    path = _node_path(given, BODY_EMAIL)
    path.write_text(json.dumps(json.loads(path.read_text()), indent=4))
    return given


def _verdicts(results: dict) -> dict:
    return {e: (r.passed if r is not None else None) for e, r in results.items()}


# --- the real captures -----------------------------------------------------

@pytest.mark.parametrize("capture", CAPTURES)
def test_real_captures_match_their_corpus(capture):
    ref = check_reference(capture)
    assert ref.identical and ref.keys_changed == []
    assert score(regrade(capture))[:2] == SCORES[capture]


@pytest.mark.parametrize("capture", CAPTURES)
def test_frozen_corpus_hashes_to_the_manifest(capture):
    manifest, _ = load_capture(capture)
    assert len(list(Path(capture, "corpus", "nodes").glob("*.json"))) == 15
    assert _corpus_hash(f"{capture}/corpus") == manifest["corpus_hash"]


@pytest.mark.parametrize("capture", CAPTURES)
def test_frozen_corpus_grades_like_the_default(capture):
    ref = check_reference(capture, f"{capture}/corpus")
    assert ref.identical
    assert _verdicts(regrade(capture, f"{capture}/corpus")) == _verdicts(regrade(capture))


def _get_email_results(capture):
    """Every get_email result the model received, parsed from the raw CLI streams."""
    for path in sorted(Path(capture, "raw").glob("day_*.jsonl")):
        tool_names = {}
        for line in path.read_text().splitlines():
            try:
                ev = json.loads(line)
            except json.JSONDecodeError:
                continue
            content = (ev.get("message") or {}).get("content")
            if not isinstance(content, list):
                continue
            for block in content:
                if not isinstance(block, dict):
                    continue
                if ev.get("type") == "assistant" and block.get("type") == "tool_use":
                    tool_names[block["id"]] = block["name"]
                elif (ev.get("type") == "user" and block.get("type") == "tool_result"
                      and tool_names.get(block.get("tool_use_id"), "").endswith("__get_email")):
                    text = block["content"]
                    if isinstance(text, list):
                        text = "".join(b.get("text", "") for b in text if b.get("type") == "text")
                    yield json.loads(text)


@pytest.mark.parametrize("capture", CAPTURES)
def test_served_email_reproduces_what_the_model_read(capture):
    """The frozen corpus, rendered by the runner's own function, gives back every email
    the model fetched, field for field."""
    manifest, _ = load_capture(capture)
    frozen = load_corpus(f"{capture}/corpus")
    plan = _plan(frozen, manifest)
    seen = list(_get_email_results(capture))
    assert len(seen) >= GET_EMAIL_FLOOR[capture]
    for got in seen:
        assert _served_email(frozen, plan, got["email_id"]) == got


# --- accepted: only the answer keys differ ---------------------------------

def _rename(email, node):
    email["answer"]["ops"][0]["create"] = "vendor shortlist"


def _rematch(email, node):
    email["answer"]["ops"][0]["match"] = ["memo"]


@pytest.mark.parametrize("capture", CAPTURES)
@pytest.mark.parametrize("change,flips", [(_rename, {KEY_EMAIL}), (_rematch, set())])
def test_key_only_change_is_scored_against_the_frozen_run(tmp_path, capture, change, flips):
    given = _frozen_copy(tmp_path, capture)
    _edit(given, KEY_EMAIL, change)
    assert _corpus_hash(str(given)) != load_capture(capture)[0]["corpus_hash"]

    ref = check_reference(capture, str(given))
    assert not ref.identical and ref.keys_changed == [KEY_EMAIL]

    before, after = _verdicts(regrade(capture)), _verdicts(regrade(capture, str(given)))
    assert {e for e in before if before[e] != after[e]} == flips
    for e in flips:                              # the grader ignores `match`; a renamed
        assert before[e] and not after[e]        # obligation no longer finds the model's object


def test_key_change_on_a_date_edge_email_that_leaves_the_plan_alone(tmp_path):
    """Decision 1: a hash check would refuse this. The deadline moves ten days, but the
    rebuilt plan is the one the run used, so it is a key repair, not a plan repair."""
    eid = "Company-Retreat.athelete-visit"
    given = _frozen_copy(tmp_path)
    _edit(given, eid, lambda e, n: e["answer"]["ops"][0]["on"].update(
        by=e["answer"]["ops"][0]["on"]["by"] + " -10d"))
    manifest, _ = load_capture(C1C)
    assert (_plan(load_corpus(given), manifest).deadlines[eid]
            != _plan(load_corpus(f"{C1C}/corpus"), manifest).deadlines[eid])

    ref = check_reference(C1C, str(given))
    assert ref.keys_changed == [eid]


def test_whitespace_reformat_is_accepted(tmp_path):
    given = _reformatted(tmp_path)
    assert _corpus_hash(str(given)) != load_capture(C1C)[0]["corpus_hash"]
    ref = check_reference(C1C, str(given))
    assert not ref.identical and ref.keys_changed == []
    assert score(regrade(C1C, str(given))) == score(regrade(C1C))


# --- refused: what the model saw, or when, would change ---------------------

def test_body_change_is_refused(tmp_path):
    given = _frozen_copy(tmp_path)
    _edit(given, BODY_EMAIL, lambda e, n: e.update(body=e["body"].replace("Hello Sir", "Hello Madam")))
    with pytest.raises(ReferenceMismatch, match=rf"{BODY_EMAIL}: body\n") as exc:
        check_reference(C1C, str(given))
    assert str(exc.value).endswith(f"pass --corpus {C1C}/corpus to grade against the run's own corpus")
    with pytest.raises(ReferenceMismatch):
        regrade(C1C, str(given))                 # library callers are protected too


@pytest.mark.parametrize("raw_field,value,served_field", [
    ("subject", "Something else entirely", "subject"),
    ("from", "CFO", "sender"),
    ("to", "CFO", "recipients"),
])
def test_header_change_is_refused(tmp_path, raw_field, value, served_field):
    given = _frozen_copy(tmp_path)
    _edit(given, BODY_EMAIL, lambda e, n: e.update({raw_field: value}))
    with pytest.raises(ReferenceMismatch, match=rf"{BODY_EMAIL}: {served_field}\n"):
        check_reference(C1C, str(given))


def test_key_change_that_moves_the_plan_is_refused(tmp_path):
    """Answer predicates feed the scheduler's deadlines, so this key edit is a plan
    repair: it needs a new run."""
    given = _frozen_copy(tmp_path)
    _edit(given, PLAN_EMAIL, lambda e, n: e["answer"]["ops"][0]["on"].update(
        eq=e["answer"]["ops"][0]["on"]["eq"] + " -20d"))
    manifest, _ = load_capture(C1C)
    moved = _plan(load_corpus(given), manifest).serve_date[PLAN_EMAIL]
    assert str(moved) != manifest["serve_plan"][PLAN_EMAIL]

    with pytest.raises(ReferenceMismatch, match="changes the serve plan") as exc:
        check_reference(C1C, str(given))
    assert (f"{PLAN_EMAIL}: served {manifest['serve_plan'][PLAN_EMAIL]} in the run, "
            f"{moved} from this corpus") in str(exc.value)


def test_key_change_that_only_reorders_a_day_is_refused(tmp_path):
    """Every serve date stays put, but the edited email becomes urgent on its day and is
    listed first. The model would have seen that day's inbox in a different order."""
    eid = "project_atlas.board-demo"
    given = _frozen_copy(tmp_path)
    _edit(given, eid, lambda e, n: e["answer"]["ops"][0]["on"].update(
        eq=e["answer"]["ops"][0]["on"]["eq"] + " -1d"))
    manifest, days = load_capture(C1C)
    plan = _plan(load_corpus(given), manifest)
    assert {e: str(d) for e, d in plan.serve_date.items()} == manifest["serve_plan"]

    day = next(rec for rec in days if eid in rec["batch"])
    with pytest.raises(ReferenceMismatch, match="changes the serve plan") as exc:
        check_reference(C1C, str(given))
    assert f"day {day['day']} ({day['serve_date']}): the run served {day['batch']}" in str(exc.value)


def test_a_day_the_run_did_not_serve_is_refused(tmp_path):
    capture = _capture_copy(tmp_path)
    (capture / "days" / "010.json").unlink()
    with pytest.raises(ReferenceMismatch, match=r"\d{4}-\d\d-\d\d: the run served nothing"):
        check_reference(str(capture), str(_reformatted(tmp_path)))


def test_corpus_whose_plan_does_not_build_is_refused(tmp_path):
    eid = "project_atlas.beta"
    given = _frozen_copy(tmp_path)
    _edit(given, eid, lambda e, n: e["answer"]["ops"][0]["on"].update(
        by=e["answer"]["ops"][0]["on"]["by"] + " -20d"))
    with pytest.raises(ReferenceMismatch, match=rf"serve plan does not build: .*{eid}"):
        check_reference(C1C, str(given))


def test_removed_email_is_refused(tmp_path):
    given = _frozen_copy(tmp_path)
    _edit(given, LEAF_EMAIL, lambda e, n: n["emails"].remove(e))
    with pytest.raises(ReferenceMismatch, match=rf"{LEAF_EMAIL}: removed \(captured\)"):
        check_reference(C1C, str(given))


def test_added_email_is_refused(tmp_path):
    given = _frozen_copy(tmp_path)
    added = LEAF_EMAIL + "-2"
    _edit(given, LEAF_EMAIL, lambda e, n: n["emails"].append({**e, "id": added}))
    with pytest.raises(ReferenceMismatch, match=rf"{added}: added"):
        check_reference(C1C, str(given))


def test_tampered_snapshot_is_refused(tmp_path):
    capture = _capture_copy(tmp_path)
    snap = _node_path(capture / "corpus", BODY_EMAIL)
    snap.write_text(snap.read_text().replace("Hello Sir", "Hello Sirs", 1))
    given = _reformatted(tmp_path)
    for corpus_dir in (given, capture / "corpus"):
        with pytest.raises(ReferenceMismatch, match="frozen corpus does not match the run"):
            check_reference(str(capture), str(corpus_dir))


def test_capture_without_a_snapshot_cannot_take_a_changed_corpus(tmp_path):
    capture = _capture_copy(tmp_path, snapshot=False)
    assert check_reference(str(capture)).identical       # the pre-C-11 path still works
    with pytest.raises(ReferenceMismatch, match="cannot verify: no frozen corpus"):
        check_reference(str(capture), str(_reformatted(tmp_path)))


# --- the CLI ---------------------------------------------------------------

def _main(monkeypatch, *argv):
    monkeypatch.setattr("sys.argv", ["sb.regrade", *argv])
    main()


def test_cli_refuses_with_exit_code_2(tmp_path, monkeypatch, capsys):
    given = _frozen_copy(tmp_path)
    _edit(given, BODY_EMAIL, lambda e, n: e.update(subject="Something else entirely"))
    with pytest.raises(SystemExit) as exc:
        _main(monkeypatch, C1C, "--corpus", str(given))
    assert exc.value.code == 2
    out, err = capsys.readouterr()
    assert out == "" and err.startswith("REFUSED")
    assert f"{BODY_EMAIL}: subject" in err
    with pytest.raises(SystemExit):
        _main(monkeypatch, C1C, "--corpus", str(given), "--json")


def test_cli_names_the_changed_keys(tmp_path, monkeypatch, capsys):
    given = _frozen_copy(tmp_path)
    _edit(given, KEY_EMAIL, _rename)
    _main(monkeypatch, C1C, "--corpus", str(given))
    out = capsys.readouterr().out
    assert (f"(run: 03e0d963b9866d8f; served emails and plan identical; "
            f"answer keys differ on 1 email)\n          {KEY_EMAIL}\n") in out
    assert f"RESCORE   {score(regrade(C1C, str(given)))[0]}/167" in out


def test_cli_on_the_run_corpus(monkeypatch, capsys):
    _main(monkeypatch, C1C, "--corpus", f"{C1C}/corpus")
    out = capsys.readouterr().out
    assert f"corpus    {C1C}/corpus  sha 03e0d963b9866d8f  (same as the run)" in out
    assert f"RESCORE   {SCORES[C1C][0]}/167" in out
