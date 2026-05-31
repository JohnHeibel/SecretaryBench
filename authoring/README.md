# SecretaryBench Scenario Editor

A visual, no-grammar editor that lets non-technical authors build benchmark
scenarios. It reads and writes the **same `corpus/nodes/*.json` files** the
benchmark runs — there is no database and no export step. Authored scenarios
drop straight into `./run.sh`.

## Run it

```bash
./run.sh ui          # builds the web app on first run, serves at http://localhost:8099
./run.sh ui-dev      # hot-reload dev mode (Vite :5173 + API :8099)
```

That's the whole thing — the backend serves the built front end and the API.

## What the author sees (and never sees)

The author works on a **graph canvas**:

- **Email cards** grouped/colored by **thread** (a thread = one `nodes/*.json`
  file with a shared cast). Click a card to open the editor drawer.
- **Arrows between cards** = "comes after". Draw one by dragging from a card's
  bottom handle to another's top handle. The editor labels each arrow
  automatically: 🔗 *after* (ordering only) or ⏰ *deadline* (a date is tied to
  an upstream email). The author never picks "static vs date" — it's inferred.

Inside an email:

- **Body** is plain prose with **date chips** dragged or clicked in from a small
  palette (`📅 Next weekday`, `📅 When it arrives`, `📌 A saved date`, …). Every
  chip shows a live preview of the actual date, and a panel shows **exactly what
  the model will read** once tokens are rendered to natural language.
- **📌 Pins (anchors).** Tick "remember this date for later emails" on any body
  chip and name it (e.g. `signing`). Downstream emails that have a path back to
  it can then pick `📌 signing` as a base — "2 weeks after 📌 signing". The pin
  dropdown only offers anchors that actually reach the current email, and
  picking one auto-creates the ⏰ deadline edge.
- **Correct answer** is built from the same chips: action (event / to-do),
  title keywords, when (exactly on / by / within / any of), length, count,
  tolerance.

All of this compiles to the governed date grammar behind the scenes — the
author types no tokens (except the deliberate "advanced / raw" escape hatch).

## Live validation

Every preview is computed by the **real engine** (`sb.resolver` /
`sb.schema` / `sb.scheduler` / `sb.oracle`):

- date chips resolve against a representative serve plan, so `@signing+2w`
  shows a real date;
- **Check solvable** runs the perfect-secretary oracle over the whole corpus —
  green means a flawless secretary *can* score 100% (the answer keys are
  satisfiable); red surfaces an unsatisfiable key or an infeasible schedule;
- the lint banner reports any structural problem (cycles, an answer that uses a
  pin with no deadline edge, an unparseable date) as the author edits.

## Architecture

```
authoring/
  chips.py        chip <-> grammar compiler (bidirectional, round-trip tested)
  corpus_io.py    tolerant corpus<->form loader/writer (never hard-fails on read)
  server.py       FastAPI: /api/corpus /resolve /render /validate /oracle /thread
  web/            React + Vite + Tailwind + React Flow front end
```

Auth is stubbed for now (single shared workspace). Set `AUTH_TOKEN` to require
an `X-Auth-Token` header; real per-user accounts slot into `server.py:_auth`
without touching the editor. The corpus directory is `corpus/` by default
(override with `CORPUS_DIR`).
