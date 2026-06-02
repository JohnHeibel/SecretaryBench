# SecretaryBench — Corpus Authoring Web App

A hosted authoring tool so club members can write benchmark emails **without
breaking the grammar or the grading criteria**. The old Excel workflow let the
email body and the answer key drift apart (the "C19" bug) and people misread how
emails depend on each other. This app removes both failure modes:

- **Scratch-style date-token blocks** (Blockly) — you snap blocks together instead
  of typing `{nth:3,FRI,+1m}`. The exact same blocks build the email body *and* the
  answer-key dates, so they can never disagree.
- **Live, real-grader preview** — every token resolves through the actual
  `sb.resolver` (vendored, not reimplemented), so what you see is what gets graded.
- **A DAG canvas** — see every email, its `static` / `date` dependencies, and the
  `@anchors` it publishes, so the structure is legible.
- **Live validation status** — the whole corpus is linted by the real
  `sb.schema.build_corpus`; the status bar is green only when it would load in the
  benchmark. Today this is an author-facing warning, not a server-side save/export block.

The app does authoring + validation + export. Running the benchmark stays local:
**Export corpus** downloads `nodes/*.json` byte-identical to what
`sb.schema.load_corpus` globs — drop them into `corpus/` and run the usual harness.

## Architecture

| Layer | Choice | Why |
|-------|--------|-----|
| Frontend | Next.js (App Router, React, TS) + Tailwind | one-click Vercel deploy |
| Block editor | Blockly (the engine behind Scratch), Zelos renderer | the "Scratch-like" ask |
| DAG | React Flow (`@xyflow/react`) | visualize emails + typed edges |
| Validation | Python serverless fns in `api/` that import the **real** `sb` | zero drift with the grader |
| Store | Neon Postgres (Vercel) — one `nodes` table, JSONB per node | corpus is small + document-shaped |

`api/resolve.py`, `api/lint.py`, and `api/oracle.py` import `sb/{resolver,schema,...}.py`, vendored into
`api/_lib/sb/` by `scripts/vendor_sb.py` (a pure copy — the anti-drift guarantee).

## Local development

```bash
cd webapp
npm install
npm run dev        # vendors sb, starts the Python validator (:8090), runs next dev (:3000)
```

No database needed locally: the store falls back to a JSON file under `.data/`,
seeded from `../corpus/nodes/`. Open http://localhost:3000.

> Plain `next dev` won't have the validator. Use `npm run dev` (it launches both),
> or run `npm run validator` and `npm run dev:next` in two terminals.

## Deploy to Vercel

1. Import the repo, set **Root Directory = `webapp`**.
2. Add the **Neon** integration (Storage tab) — it sets `POSTGRES_URL` automatically.
3. Set `APP_PASSCODE` to a shared club passcode for the current login flow. This is
   not full route/API enforcement yet; omit for open local/dev use.
4. Deploy. `npm run build` vendors `sb` first; `requirements.txt` enables the Python
   functions. They serve `/api/resolve`, `/api/lint`, and `/api/oracle` at the same origin (no
   `NEXT_PUBLIC_VALIDATOR_BASE` needed in prod).

## Round-trip with the benchmark

Export → unzip into `corpus/` → `python -m sb.demo` (oracle) should solve it 100%.
That oracle pass is the proof the app-authored corpus is valid and drift-free.
