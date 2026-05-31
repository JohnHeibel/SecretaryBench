"""FIX-10: the stress-chain generator produces a loadable single-scenario chain
long enough to push claude past its context window.

Uses a small config for speed; the real compaction run uses the script's
defaults (60 emails x ~12k chars). No server needed — this only checks the
generated workbook parses correctly."""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from tests.generate_stress_chain import generate
from loader import load_scenarios


def test_generated_chain_loads_as_one_scenario(tmp_path):
    path = str(tmp_path / "stress.xlsx")
    generate(path, n_emails=12, body_chars=800)

    scenarios = load_scenarios(path)
    assert len(scenarios) == 1, scenarios
    s = scenarios[0]
    assert len(s.emails) == 12
    # chain order preserved
    assert [e.email_number for e in s.emails] == list(range(1, 13))
    # only the final email carries a checkable criterion, so the chain scores
    assert s.emails[-1].success_criteria == "CC-{date}"
    assert s.success_criteria == ["CC-{date}"]


def test_body_size_scales_with_param(tmp_path):
    path = str(tmp_path / "stress2.xlsx")
    generate(path, n_emails=5, body_chars=4000)
    s = load_scenarios(path)[0]
    # every body is sized near the requested target (realistic bulk, not lorem)
    for e in s.emails:
        assert 3000 <= len(e.body) <= 4200, len(e.body)
