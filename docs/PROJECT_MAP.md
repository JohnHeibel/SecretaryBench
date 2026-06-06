# SecretaryBench project map

This is the quick orientation map. It is not a replacement for the design docs;
it tells you where things live and what each part is responsible for.

## Main specs and history

- `HOW_IT_WORKS.md`: the plain-language tour of the whole machine + the worked Project Atlas example. Read first.
- `GRADING_MODEL.md`: the precise grading contract (scoped, binary, per email).
- `ANSWER_KEY_GRAMMAR.md`: the token grammar and verb-based answer key contract.
- `AUTHORING_GUIDE.md`: legacy low-level authoring pattern guide (type mix, tiers, needles); use
  `HOW_IT_WORKS.md`, `docs/AUTHORING_WALKTHROUGH.md`, and `/guide` for the current webapp flow.
- `docs/AUTHORING_WALKTHROUGH.md`: click-by-click "author one storyline" example (focus mode, CEO-sent emails).
- `BACKLOG.md`: deferred features (ranged offsets; global calendar + grading, incl. cross-storyline time conflicts §2a) + triggers to revisit.
- `OPEN_QUESTIONS.md`: design calls still being decided.
- `RUNNING.md`: plain-language guide for building a corpus, running a model, and
  analyzing results.
- `TIER_LIST.md`: authoring playbook for T1/T2/T3 email difficulty.
- `RUN_RESULTS.md`: historical Haiku retrieval-span pilot results.
- `docs/POTENTIAL_GAMING.md`: monitor-first notes for broad search and suspicious
  `email_id` attribution risks.
- `DAY_LOOP_DESIGN_ISSUE.md`: historical design issue from before the day loop was
  rebuilt (resolved; kept as decision-record context for `adr/0001`).
- `docs/design/BENCHMARK_REDESIGN.md`: why the project moved away from Excel and
  toward a governed DAG benchmark.
- `docs/design/SERVING_AND_SCHEMA.md`: current serving model, edge types, and JSON
  corpus schema.
- `docs/history/dag-transition-2026-06-02.md`: short transition note from the
  Excel pipeline to the DAG pipeline.
- `docs/adr/0001-obligations-replace-count.md`: decision record for replacing
  count-style grading with named verb obligations.
- `docs/MCP.md`: MCP setup notes.
- `docs/api_reference.md`: legacy/internal API reference kept for context.

## Corpus and core engine

- `corpus/nodes/*.json`: authored benchmark nodes. Each node file contains emails,
  date tokens, dependency edges, and answer-key ops.
- `sb/schema.py`: JSON data model, loader, linter, DAG validation, node sugar
  expansion, anchor reachability checks, and answer-key wiring.
- `sb/resolver.py`: grammar evaluator. It renders date tokens for emails and
  resolves the same expressions for grading.
- `sb/scheduler.py`: seeded DAG scheduler. It serves emails by day, respects
  `static` and `date` edges, computes date-edge deadlines, and produces the run plan.
- `sb/engine.py`: deterministic in-memory serve/grade loop used by tests and the
  oracle.
- `sb/grader.py`: state-based grader for `create`, `move`, `cancel`, and no-action
  emails.
- `sb/oracle.py`: perfect reference model that reads the answer key directly. A
  valid corpus should be oracle-solvable at 100%.
- `sb/span.py`: computes how far back a needed anchor fact sits in the served email
  stream.
- `sb/scale.py`: copies the authored corpus into `build/scaled` and injects
  no-action filler to increase retrieval span.
- `sb/analyze.py`: reads a live run log and reports accuracy by tier, span, and
  inbox-search behavior.
- `sb/demo.py`: small demo entry point for local experimentation.
- `sb/sync.py`: sync helpers between persisted node data and JSON files.

## Live model runner

- `run.sh`: convenience wrapper for live runs.
- `sb/live/runner.py`: runs a real model one simulated day at a time, drops that
  day's mail into the store, invokes the model with MCP tools, snapshots state, and
  grades the day's emails.
- `sb/live/store_app.py`: FastAPI in-memory live store for calendars, todos, inbox,
  current day, and run state.
- `sb/live/mcp_app.py`: MCP tool surface exposed to the model, including inbox
  listing, email reading, search, and calendar/todo actions.

## Tests and fixtures

- `sb/tests/test_schema.py`: schema loader and linter coverage.
- `sb/tests/test_resolver.py`: date grammar and rendering coverage.
- `sb/tests/test_scheduler.py`: DAG scheduling, deadlines, reproducibility, and
  feasibility coverage.
- `sb/tests/test_e2e.py`: load, plan, serve, and grade tests with perfect and
  imperfect mock models.
- `sb/tests/fixtures/nodes/*.json`: small hand-authored test corpus.

## Web authoring app

- `webapp/package.json`: Next.js app scripts and dependencies.
- `webapp/app/page.tsx`: main authoring workspace route.
- `webapp/app/guide/page.tsx`: in-app authoring guide.
- `webapp/app/login/page.tsx`: lightweight login page.
- `webapp/app/api/nodes/*`: JSON node persistence API.
- `webapp/app/api/export/route.ts`: corpus export API.
- `webapp/app/api/auth/route.ts`: simple auth endpoint.
- `webapp/api/lint.py`: Python lint endpoint using the real `sb.schema` code.
- `webapp/api/oracle.py`: Python oracle endpoint using the real core engine.
- `webapp/api/resolve.py`: Python resolver endpoint for token preview.
- `webapp/api/_lib/apihelp.py`: shared helpers for Python API functions.
- `webapp/scripts/vendor_sb.py`: vendors the pure-Python `sb` modules into the
  webapp API bundle and generates TypeScript schema constants.
- `webapp/scripts/validator_server.py`: local Python validator service for dev.
- `webapp/scripts/dev.mjs`: starts the webapp plus validator during development.
- `webapp/lib/types.ts`: TypeScript app types for nodes, emails, edges, and answers.
- `webapp/lib/store.ts`: client-side state management.
- `webapp/lib/grammar.ts`: client grammar helpers for authoring UI.
- `webapp/lib/schema.generated.ts`: generated constants from `sb.schema`.
- `webapp/lib/api.ts`: client API helpers.
- `webapp/components/Workspace.tsx`: main authoring surface.
- `webapp/components/Sidebar.tsx`: node/email navigation.
- `webapp/components/EmailEditor.tsx`: email metadata and body editor.
- `webapp/components/BodyEditor.tsx`: token-aware body editing.
- `webapp/components/AnswerKeyBuilder.tsx`: verb-op answer-key editor.
- `webapp/components/DagCanvas.tsx`: dependency graph view.
- `webapp/components/ValidateBar.tsx`: lint/oracle validation controls.
