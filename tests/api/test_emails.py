from datetime import datetime, timezone
from fastapi.testclient import TestClient

from app.main import app
from app import store

client = TestClient(app)

BASE_DATE = datetime(2026, 4, 15, 10, 0, 0, tzinfo=timezone.utc).isoformat()

SAMPLE_EMAIL = {
    "email_id": 1,
    "subject": "Hello",
    "created_at": BASE_DATE,
    "sender": "alice@example.com",
    "recipients": ["bob@example.com"],
    "body": "Hi there",
}

SAMPLE_SCENARIO = {
    "scenario_id": 1,
    "emails": [SAMPLE_EMAIL],
}


def setup_function():
    store.emails.clear()
    store.scenarios.clear()


# --- List ---

def test_list_emails_empty():
    response = client.get("/emails/")
    assert response.status_code == 200
    assert response.json() == []


def test_list_emails_returns_all():
    client.post("/scenarios/", json=SAMPLE_SCENARIO)
    response = client.get("/emails/")
    assert response.status_code == 200
    assert len(response.json()) == 1
    assert response.json()[0]["email_id"] == 1


# --- Get ---

def test_get_email_returns_200():
    client.post("/scenarios/", json=SAMPLE_SCENARIO)
    response = client.get("/emails/1")
    assert response.status_code == 200
    assert response.json()["email_id"] == 1
    assert response.json()["subject"] == "Hello"


def test_get_email_not_found_returns_404():
    response = client.get("/emails/999")
    assert response.status_code == 404


# --- Delete (intentionally unsupported) ---
# DELETE /emails was deliberately removed: the email log is append-only across a
# run (README.md). The route does not exist, so the API returns 405. This test
# replaces four stale tests that asserted a 204/404 delete contract.

def test_delete_email_not_allowed():
    client.post("/scenarios/", json=SAMPLE_SCENARIO)
    response = client.delete("/emails/1")
    assert response.status_code == 405, (
        "DELETE /emails is intentionally unsupported (append-only log); "
        f"got {response.status_code}"
    )
