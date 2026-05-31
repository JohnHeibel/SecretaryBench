"""FIX-3 (token/tool/compaction logging + message-id dedup) and FIX-9
(per-scenario compaction dimension) on the shared core parser.

No server or claude needed — these drive harness.base.parse_stream_output with
recorded stream-json and inspect the side effects."""
from __future__ import annotations

import json
import os

from harness import base, get_adapter

_FIXTURE = os.path.join(os.path.dirname(__file__), "fixtures", "sample_stream_json.jsonl")


# --- FIX-3: dedup + token/tool logging -------------------------------------

def test_dedup_and_token_tool_logging(tmp_path, monkeypatch):
    monkeypatch.setattr(base, "TOKEN_LOG_PATH", str(tmp_path / "tok.jsonl"))
    monkeypatch.setattr(base, "TOOL_LOG_PATH", str(tmp_path / "tool.jsonl"))
    base.release_scenario(777)

    with open(_FIXTURE) as f:
        sample = f.read()

    sid = base.parse_stream_output(sample, scenario_id=777, email_index=1)
    assert sid == "sess-abc"

    tel = base.scenario_telemetry(777)
    # msg_1 is emitted twice (pre- and post-tool_use). Dedup keeps it ONCE, so
    # only two rounds are counted (msg_1 + msg_2), not three.
    assert tel["rounds"] == 2, tel
    # total input = (200+1000+0) + (50+0+1200) = 1200 + 1250 = 2450
    assert tel["input_tokens"] == 2450, tel
    assert tel["output_tokens"] == 23, tel  # 15 + 8

    # One token-log row per counted round.
    tok_rows = [json.loads(l) for l in open(tmp_path / "tok.jsonl") if l.strip()]
    assert len(tok_rows) == 2, tok_rows
    # Cache creation vs read are split (proves caching is observable).
    assert any(r["cache_creation_input_tokens"] == 1000 for r in tok_rows)
    assert any(r["cache_read_input_tokens"] == 1200 for r in tok_rows)

    # The deduped emission's tool_use is logged exactly once.
    tool_rows = [json.loads(l) for l in open(tmp_path / "tool.jsonl") if l.strip()]
    assert len(tool_rows) == 1, tool_rows
    assert tool_rows[0]["tool_name"] == "create_event"


def test_result_error_raises():
    stream = "\n".join([
        '{"type":"system","subtype":"init","session_id":"s"}',
        '{"type":"result","session_id":"s","is_error":true,"result":"boom"}',
    ])
    try:
        base.parse_stream_output(stream, scenario_id=1, email_index=1)
        assert False, "expected RuntimeError on is_error result"
    except RuntimeError as e:
        assert "boom" in str(e) or "failed" in str(e)


# --- FIX-9: compaction attributed per scenario -----------------------------

def test_compaction_event_attributed_per_scenario(monkeypatch, tmp_path):
    monkeypatch.setattr(base, "TOKEN_LOG_PATH", "")  # don't pollute cwd
    base.release_scenario(888)
    stream = "\n".join([
        '{"type":"system","subtype":"init","session_id":"s"}',
        '{"type":"system","subtype":"compaction"}',
        '{"type":"assistant","message":{"id":"a","usage":{"input_tokens":10,"output_tokens":2},"content":[]}}',
        '{"type":"result","session_id":"s"}',
    ])
    base.parse_stream_output(stream, scenario_id=888, email_index=1)
    tel = base.scenario_telemetry(888)
    assert tel["compactions"] == 1, tel


def test_compaction_via_context_management_field(monkeypatch):
    monkeypatch.setattr(base, "TOKEN_LOG_PATH", "")
    base.release_scenario(889)
    stream = "\n".join([
        '{"type":"system","subtype":"init","session_id":"s"}',
        '{"type":"assistant","message":{"id":"a","context_management":{"compacted":true},"usage":{"input_tokens":5,"output_tokens":1},"content":[]}}',
        '{"type":"result","session_id":"s"}',
    ])
    base.parse_stream_output(stream, scenario_id=889, email_index=1)
    assert base.scenario_telemetry(889)["compactions"] == 1


def test_adapter_exposes_scenario_telemetry(monkeypatch):
    monkeypatch.setattr(base, "TOKEN_LOG_PATH", "")
    base.release_scenario(999)
    base.parse_stream_output(
        '{"type":"system","subtype":"compaction"}', scenario_id=999, email_index=1)
    a = get_adapter("claude-code", model="m", reasoning=None, api_base=None,
                    openrouter_key=None, conversation_continuity=None)
    assert a.scenario_telemetry(999)["compactions"] == 1


def _turn(msg_id: str, cache_read: int) -> str:
    return "\n".join([
        '{"type":"system","subtype":"init","session_id":"s"}',
        ('{"type":"assistant","message":{"id":"%s","usage":{"input_tokens":10,'
         '"cache_read_input_tokens":%d,"output_tokens":2},"content":[]}}' % (msg_id, cache_read)),
        '{"type":"result","session_id":"s"}',
    ])


def test_silent_compaction_detected_via_context_drop(monkeypatch):
    # Claude Code 2.1.158 print mode compacts silently — no system/compaction
    # event. We detect it by the context drop it leaves (193k -> 18k).
    monkeypatch.setattr(base, "TOKEN_LOG_PATH", "")
    base.release_scenario(555)
    base.parse_stream_output(_turn("m1", 193_000), scenario_id=555, email_index=26)
    assert base.scenario_telemetry(555)["compactions"] == 0  # growth so far, no drop
    base.parse_stream_output(_turn("m2", 18_000), scenario_id=555, email_index=27)
    assert base.scenario_telemetry(555)["compactions"] == 1  # the drop = compaction


def test_no_false_compaction_on_normal_growth(monkeypatch):
    monkeypatch.setattr(base, "TOKEN_LOG_PATH", "")
    base.release_scenario(556)
    for i, tok in enumerate([60_000, 70_000, 80_000], start=1):
        base.parse_stream_output(_turn(f"m{i}", tok), scenario_id=556, email_index=i)
    assert base.scenario_telemetry(556)["compactions"] == 0


def test_small_context_drop_not_flagged(monkeypatch):
    # A drop from a small prior context (< 50k) is normal variation, not compaction.
    monkeypatch.setattr(base, "TOKEN_LOG_PATH", "")
    base.release_scenario(558)
    base.parse_stream_output(_turn("m1", 40_000), scenario_id=558, email_index=1)
    base.parse_stream_output(_turn("m2", 4_000), scenario_id=558, email_index=2)
    assert base.scenario_telemetry(558)["compactions"] == 0


def test_release_scenario_clears_telemetry(monkeypatch):
    monkeypatch.setattr(base, "TOKEN_LOG_PATH", "")
    base.parse_stream_output(
        '{"type":"system","subtype":"compaction"}', scenario_id=1234, email_index=1)
    assert base.scenario_telemetry(1234)["compactions"] == 1
    base.release_scenario(1234)
    assert base.scenario_telemetry(1234)["compactions"] == 0


# --- bench_logger.compaction_breakdown -------------------------------------

def test_compaction_breakdown_silent_when_empty(capsys):
    import bench_logger as blog
    empty = {"within_window": {"count": 0, "score": 0, "max_score": 0},
             "through_compaction": {"count": 0, "score": 0, "max_score": 0}}
    blog.compaction_breakdown(empty, 0, 0, 0)
    assert capsys.readouterr().out == ""


def test_compaction_breakdown_prints_when_data(capsys):
    import bench_logger as blog
    split = {"within_window": {"count": 2, "score": 1, "max_score": 2},
             "through_compaction": {"count": 1, "score": 1, "max_score": 1}}
    blog.compaction_breakdown(split, 1, 5000, 200, ungraded=3)
    out = capsys.readouterr().out
    assert "Compaction" in out
    assert "within window" in out and "through compaction" in out
