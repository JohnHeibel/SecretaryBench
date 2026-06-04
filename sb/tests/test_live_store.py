from fastapi.testclient import TestClient

from sb.live.store_app import app


def test_store_logs_potential_gaming_warnings():
    client = TestClient(app)
    client.post("/reset")
    client.post("/day", json={"date": "2026-06-01"})
    client.post("/inbox", json={
        "email_id": "n.today", "sender": "sender", "recipients": ["CEO"],
        "subject": "today", "body": "body", "served_date": "2026-06-01",
    })

    client.get("/inbox/search", params={"limit": 0})
    client.post("/events", json={
        "email_id": "fake.id", "title": "hidden event",
        "start": "2026-06-02T09:00:00", "end": "2026-06-02T10:00:00",
    })

    kinds = [w["kind"] for w in client.get("/warnings").json()]
    assert "broad_search" in kinds
    assert "large_search_result" in kinds
    assert "invalid_email_id" in kinds


def test_store_logs_stale_email_id_warning():
    client = TestClient(app)
    client.post("/reset")
    client.post("/day", json={"date": "2026-06-01"})
    client.post("/inbox", json={
        "email_id": "n.old", "sender": "sender", "recipients": ["CEO"],
        "subject": "old", "body": "body", "served_date": "2026-06-01",
    })
    client.post("/day", json={"date": "2026-06-02"})
    client.post("/todos", json={"email_id": "n.old", "title": "stale todo", "due_date": "2026-06-03T17:00:00"})

    assert any(w["kind"] == "stale_email_id" for w in client.get("/warnings").json())
