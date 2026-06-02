from sb.analyze import parse_log


def test_parse_log_counts_search_inbox_at_day_level(tmp_path):
    log = tmp_path / "run.log"
    log.write_text("""
── day 1 · Mon Jun 01 · 2 new email(s) ──
   tools  list_new_emails, get_email

 PASS  [1] alpha.brief  - served Mon Jun 01
 PASS  [2] alpha.review  - served Mon Jun 01

── day 2 · Tue Jun 02 · 2 new email(s) ──
   tools  list_new_emails, get_email, search_inbox, create_event

 PASS  [3] beta.kickoff  - served Tue Jun 02
 FAIL  [4] beta.followup  - served Tue Jun 02
""")

    results = parse_log(str(log))

    assert results["alpha.brief"] == {"passed": True, "searched": False}
    assert results["alpha.review"] == {"passed": True, "searched": False}
    assert results["beta.kickoff"] == {"passed": True, "searched": True}
    assert results["beta.followup"] == {"passed": False, "searched": True}
