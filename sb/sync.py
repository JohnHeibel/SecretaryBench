"""
sb.sync — pull the authored corpus down from the web app's database into
corpus/nodes/*.json, so authoring (web app + Neon) flows into a benchmark run
with one command.

The web app's Postgres/Neon DB is the live AUTHORING store; the runner reads git
JSON files (see sb.live.runner). This bridges them: it GETs /api/nodes from a
deployed (or local) web app and writes each node as corpus/nodes/<id>.json. Git
stays the reproducible source of truth for a run — this just automates the
"Export corpus -> unzip into corpus/" step.

Safety: the pulled corpus is validated IN MEMORY (the same sb.schema.build_corpus
gate the web app lints with) BEFORE anything is written, so a broken pull can never
clobber a good local corpus.

    # against a deployed app (passcode-gated):
    SB_WEBAPP_URL=https://yourapp.vercel.app APP_PASSCODE=… python -m sb.sync
    # against local dev:
    python -m sb.sync --url http://localhost:3000
    python -m sb.sync --prune        # also delete local node files not in the DB
"""
from __future__ import annotations

import argparse
import json
import os
import tempfile
from pathlib import Path

import httpx

from sb.schema import build_corpus

REPO = Path(__file__).resolve().parents[1]


def fetch_nodes(url: str, passcode: str | None) -> list[dict]:
    cookies = {"sb_pass": passcode} if passcode else None
    r = httpx.get(f"{url.rstrip('/')}/api/nodes", cookies=cookies, timeout=30, follow_redirects=False)
    if r.is_redirect or "/login" in str(r.url):
        raise SystemExit("auth required: set APP_PASSCODE (or --passcode) to match the deployed app.")
    r.raise_for_status()
    data = r.json()
    nodes = data.get("nodes", data) if isinstance(data, dict) else data
    if not isinstance(nodes, list):
        raise SystemExit(f"unexpected /api/nodes response shape: {type(nodes).__name__}")
    return nodes


def _safe_path(node_dir: Path, node_id: str) -> Path:
    """Resolve <id>.json and refuse anything that escapes node_dir (path traversal).
    build_corpus already rejects bad ids, but corpus/ is git — belt and suspenders."""
    p = (node_dir / f"{node_id}.json").resolve()
    if not p.is_relative_to(node_dir.resolve()):
        raise SystemExit(f"refusing to write node id {node_id!r}: escapes {node_dir}")
    return p


def sync(url: str, passcode: str | None, dest: Path, prune: bool, force: bool = False) -> None:
    nodes = fetch_nodes(url, passcode)
    # Validate in memory FIRST — refuse to write a corpus the runner would reject.
    corpus = build_corpus(nodes)

    node_dir = dest / "nodes"
    existing = {f.stem for f in node_dir.glob("*.json")} if node_dir.exists() else set()
    # An empty (or drastically shrunken) pull is almost always a misconfig — a wrong
    # URL, an unseeded DB, a dropped connection — NOT an intentional wipe. build_corpus([])
    # is "valid", so this guard, not validation, is what protects the local corpus.
    if not nodes:
        raise SystemExit("refusing to sync: /api/nodes returned 0 nodes (wrong URL or unseeded DB?)")
    if existing and len(nodes) < 0.5 * len(existing) and not force:
        raise SystemExit(f"refusing to sync: pull has {len(nodes)} nodes but {len(existing)} exist locally — pass --force if intended")
    print(f"pulled {len(nodes)} nodes / {len(corpus.emails)} emails — corpus is valid")

    node_dir.mkdir(parents=True, exist_ok=True)
    pulled_ids = {n["id"] for n in nodes}
    for nid in pulled_ids:  # validate every target path before touching disk
        _safe_path(node_dir, nid)

    # Write to temp files then atomically rename, so an interrupt can't leave a
    # half-written corpus that is neither the old nor the new state.
    for node in sorted(nodes, key=lambda n: n["id"]):
        final = _safe_path(node_dir, node["id"])
        fd, tmp = tempfile.mkstemp(dir=node_dir, suffix=".tmp")
        os.close(fd)
        Path(tmp).write_text(json.dumps(node, indent=2) + "\n")
        os.replace(tmp, final)
    added = pulled_ids - existing
    try:
        shown = node_dir.relative_to(REPO)
    except ValueError:
        shown = node_dir
    print(f"wrote {len(nodes)} files -> {shown}" + (f"  (+{len(added)} new: {', '.join(sorted(added))})" if added else ""))

    stale = existing - pulled_ids
    if stale:
        if prune:
            for sid in sorted(stale):
                _safe_path(node_dir, sid).unlink()
            print(f"pruned {len(stale)} local node(s) not in the DB: {', '.join(sorted(stale))}")
        else:
            print(f"note: {len(stale)} local node(s) not in the DB (kept; pass --prune to remove): {', '.join(sorted(stale))}")


def main() -> None:
    ap = argparse.ArgumentParser(description="Pull the authored corpus from the web app DB into corpus/nodes/.")
    ap.add_argument("--url", default=os.environ.get("SB_WEBAPP_URL", "http://localhost:3000"), help="web app base URL (env SB_WEBAPP_URL)")
    ap.add_argument("--passcode", default=os.environ.get("APP_PASSCODE"), help="APP_PASSCODE if the app is gated (env APP_PASSCODE)")
    ap.add_argument("--dest", default="corpus", help="corpus dir to write into (default: corpus)")
    ap.add_argument("--prune", action="store_true", help="delete local node files not present in the DB")
    ap.add_argument("--force", action="store_true", help="allow a pull much smaller than the local corpus")
    a = ap.parse_args()
    sync(a.url, a.passcode, REPO / a.dest, a.prune, a.force)


if __name__ == "__main__":
    main()
