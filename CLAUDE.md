# SecretaryBench — orientation for Claude

A map so we don't have to re-explore every session. This is the **index**; the
linked docs and the code are canonical. Treat any file:line or behavior claim here as
"probably true, verify before asserting" — code moves, this drifts.

## What this is

A benchmark that measures an AI assistant's **long-horizon temporal reasoning**: can it
manage a busy CEO's calendar over many simulated days, working only from emails in an
inbox (schedule / move / cancel events and to-dos, resolve relative dates like "two
weeks after the kickoff", handle authored reschedule conflicts)?

Two halves:
- **`sb/`** — the Python benchmark engine (the grammar, grader, scheduler, oracle, live runner). This is the source of truth for *how grading works*.
- **`webapp/`** — a Next.js authoring tool so club members write scenarios without breaking the grammar/grading. It **vendors a pure copy of `sb/`** so what you see is what gets graded.

## Repo map

```
sb/                  benchmark engine (Python)
  resolver.py        the date/time grammar: parse + evaluate an expr -> date/datetime/Interval/TimeInterval
  schema.py          corpus data model + JSON loader + linter (build_corpus / load_corpus)
  grader.py          state-based, binary, per-email grading
  scheduler.py       build_plan: assigns serve days, respects deadlines (date edges), feasibility
  engine.py          Store + run(corpus, plan, model) -> RunResult (.score())
  oracle.py          reference "perfect secretary" that acts straight from the answer key
  live/              the runnable harness (MCP app, runner) for real models
  sync.py            pull the webapp DB -> corpus/nodes/*.json ; demo.py / scale.py / analyze.py helpers
  tests/             pytest
corpus/nodes/*.json  the actual scenarios (CURRENTLY EMPTY — authored in the webapp, pulled via sb.sync)
webapp/              Next.js authoring tool (see webapp section below)
docs/                deeper notes: PROJECT_MAP.md, RUNNING.md, MCP.md, TIER_LIST.md, adr/, design/, history/
*.md (root)          the settled specs (see "Canonical docs")
```

## Mental model (read GRADING_MODEL.md / HOW_IT_WORKS.md for the full version)

- **Scenario = node = storyline**: one DAG of date-dependent emails sharing a cast (e.g. "the Q3 board saga"). Authored and verified in isolation; executed interleaved with other nodes on one shared calendar/inbox at run time.
- **Grading is binary, per email, attribution-scoped to its node.** An email scores **1** iff the model did *exactly* its authorized ops (each create/move/cancel resolves to exactly one matching object on the right day/time; cancel leaves none) **and nothing else**. Otherwise 0. No partial credit. A no-ops email (FYI/junk) must create nothing.
- **Three layers**: *spine* (immutable, precomputed dates/deadlines — model can't touch), *workspace* (the objects the model creates — the only mutable thing), *memory* = the inbox (immutable ground truth, everything recoverable via search).
- **Deterministic**: each email's grade is a pure function of (its answer key, its serve date, its DAG ancestors). The seed only shuffles *which* scenarios run, *when* emails arrive, and *how much* noise surrounds them.
- **Exact matching** (day + time). The only deliberate multi-day or bounded answers are `by` (deadline, day-level unless timed) and `any_of` (options). The old `within:Nd` tolerance is retired.
- **Conflicts are authored as reschedules** (a later email moves/cancels an earlier event), never emergent. There is deliberately **no "find an open slot" / `no_overlap`** — it had no single deterministic answer.
- **Thread fence** (engine Phase 2, not authoring): the store blocks cross-scenario edits and docks the offending email. Authors ignore it.

## The grammar (resolver.py / webapp lib/dateExpr.ts — they mirror each other)

```
expr     := base (offset)* (time)?         ( base defaults to `serve` if it starts with an offset )
base     := serve | @NAME | next:WD [from expr] | this:WD [from expr]
          | nth:(N|last),WD,monthref | dom:D,monthref | week_of:(expr) | month:monthref
offset   := (+|-) INT (d|bd|w|m|y)         ( bd = business days )
monthref := 0m | (+|-) INT m
time     := @HH:MM [ - HH:MM ]             ( static clock; work hours 05:00–23:00; end > start )
```

- Values: **date** (bare day), **datetime** (`@HH:MM` point), **Interval** (whole week/month), **TimeInterval** (`@HH:MM-HH:MM` within-day span). Bare dates grade day-level; timed exprs grade to the minute. A `by` predicate with a time compares start only against that cutoff.
- **Anchors** = reusable named dates. Published by `{!name = expr}` in a body or `answer.emits`. A later email reuses it as `@name`. An email whose **answer** reuses an ancestor's anchor is a **needle** (the long-horizon retrieval test) and needs a `date` dependency edge.

## Answer key (schema.py)

```
answer = { ops: [ ... ], emits?: {name: expr} }
op     = exactly one verb: { create|move|cancel: "<obligation name>", kind?, on?, match?, tolerance? }
         create: needs kind ("event"|"todo") + `on` predicate
         move:   needs `on`              cancel: takes neither on nor kind (inherits from the create)
on (predicate) = one of: { eq } | { by } | { in (+ optional not_in) } | { any_of: [...] }
match  = title keywords the grader matches (defaults to [name]); the MODEL never sees them
```

- An **obligation** = a named thing (its name is its identity *and* a node-scoped `@anchor`). `move`/`cancel` target it by name; their dependency edge to the creator is auto-derived by `schema.py`.
- Lint (in `build_corpus`) checks: acyclic DAG, anchors reachable from ancestors, tokens parse, needles carry a `date` edge, no keyword collisions.

## webapp/ specifics

- **Anti-drift is the whole point.** `api/{resolve,lint,oracle}.py` import the REAL `sb`, vendored to `api/_lib/sb/` by `scripts/vendor_sb.py` (pure copy; also generates `lib/schema.generated.ts` from `sb.schema`). **`python3 scripts/vendor_sb.py --check` must pass** (no drift). Never reimplement the grammar in JS — `lib/dateExpr.ts` is a *mirror* whose only job is the live builder; the Python resolver is authoritative (via `/api/resolve`).
- `/api/resolve` = live "this resolves to…" preview. `/api/lint` = whole-corpus well-formed gate. `/api/oracle` = satisfiability: `build_corpus → build_plan → engine.run(oracle_model)`, **ok iff score == 1.0** (a perfect secretary could actually do it).
- **Key UI files**: `Workspace.tsx` (state + orchestration), `Sidebar.tsx`, `EmailEditor.tsx`, `AnswerKeyBuilder.tsx`, `DateBuilder.tsx` (the inline date/time builder), `BodyEditor.tsx`, `DagCanvas.tsx`. **Key lib**: `dateExpr.ts`, `grammar.ts`, `types.ts`, `schema.generated.ts`, `templates.ts` (the loadable "Project Atlas" example — the one compact example for the headline constructs at scale; Helios was removed), `store.ts`, `api.ts`.
- **Store**: Postgres (Neon) on Vercel; locally falls back to `webapp/.data/nodes.json` (gitignored). Both seed from `webapp/seed/nodes.json` (a build-time snapshot of `corpus/nodes/`).
- **Export-only, by design.** Author in the webapp → `/api/export` (or `python -m sb.sync`) pulls JSON out → run. There is **no import and never will be**; the only inputs are handwritten emails. Don't propose one.
- **No auth in prod** — the deployed app is public on purpose. Don't re-add auth unless asked.
- **The authoring UI is deliberately minimal (2026-06 simplification).** No corpus-mix panel, no T1/T2/T3 difficulty buttons, no action/FYI/junk role control (`CompositionPanel.tsx` + `lib/composition.ts` were deleted). Restraint is a single "this email needs no action" checkbox in the answer key (ticking it sets `ops: []`). `Email.tier` / `Email.type` stay optional in `types.ts` so old data still loads, but nothing sets or shows them. The date builder's predicate menu offers eq / by / any_of (`in` / `not_in` remain in the grammar but aren't surfaced); `by` deadlines can carry a clock cutoff for events or to-dos. The base menu now offers all eight bases (serve / anchor / next / this / nth / dom / `week_of` / `month`). Depends-on is one auto-adding select (new edges default to `static`; needles auto-upgrade to `date`). All UI copy is em-dash-free and the `/guide` page is the plain-language walkthrough.
- **Exec-feedback pass (2026-06-05, `webapp-v2`) — four idiomatic/baby-proofing changes.** (1) **Date entry is builder-only**: the author-facing "type it" raw box is gone (`DateBuilder.tsx`); `week_of`/`month` were added to the base dropdown so the structured builder covers every form. The raw `RawField` now renders ONLY to display/repair a stored value that can't parse (legacy `in:`/`not_in`/exotic), never as a typing entry point. (2) **Variables scoped to the selected storyline**: `Workspace.tsx` computes `anchors` from `nodes.filter(selNode)` (was the validate set), so the `@anchor` pickers can't reference another storyline; lint/oracle still run on `focusClosure`. (3) **Delete confirms** on remove-node / remove-email (`Workspace.tsx`) and remove-cast-person (`EmailEditor.tsx`, warns if used in From/To/Cc). (4) **Standardized people** (`lib/people.ts`): `STANDARD_ROSTER` picker + "custom person" in `CastManager`; cast key normalized to `UPPER_SNAKE` (≤24 chars) with every From/To/Cc reference rewritten on rename; display name whitespace-collapsed + capped (≤40). Cast stays per-node in export (roster just materializes in).
- **Deep-link focus mode (`webapp-v2`, for ~20 concurrent authors).** `?node=<id>` scopes one author to a single storyline: `Workspace.tsx` reads it via `useSearchParams` (so `app/page.tsx` wraps `<Workspace>` in `<Suspense>`), the sidebar/DAG/depends-on picker show only that node, and **lint/oracle run on `focusClosure(nodes, id)`** (the node + any it truly references; usually just itself) instead of the whole corpus — so per-author validation is O(one scenario) and the bottom bar reflects that storyline alone. The bare site (no param) is the coordinator view: full list + per-storyline "copy link" + whole-corpus check + export. `renameNode`/`renameEmail` now save only the rows that actually changed (was: every node), so one author can't clobber another's in-flight node. Walkthrough: `docs/AUTHORING_WALKTHROUGH.md`.

## Gotchas / invariants (the non-obvious stuff)

- **Obligation `@names` resolve ONLY in answer predicates, not in body tokens** (the schema rewrites predicates, not bodies). So the webapp offers obligation names as `@anchors` only in the answer-key builder; body anchors come from `{!name=…}`.
- **A `move`/`cancel` op stores no `kind`** (the schema fills it from the create at build time). The webapp derives it via `obligationKinds(node)` so a reschedule knows it targets an event (→ offer a clock).
- A timed `@HH:MM` outside 05:00–23:00, or end ≤ start, is a **grammar error** (rejected at lint), surfaced live in the builder by `timeError`.
- **Email `from` / `to` / `cc` are inbox presentation, never graded.** `to` accepts a string OR a list (schema normalizes to `recipients: list[str]`); the webapp stores it as a multi-recipient chip picker. `cc` is an optional list (additive field, defaults `[]`, shown to the model via the live `get_email` tool, never read by the grader). Adding or changing recipients/cc cannot change a score, so the oracle stays 1.0.
- `corpus/nodes/` is empty in git; the live corpus lives in the webapp store and is pulled out via `sb.sync`. The corpus is recoverable from there / git.

## Commands (verify before declaring done)

```bash
# webapp/  (run from webapp/)
npm run dev                       # vendor + Python validator :8090 + next dev :3000
npm run typecheck                 # tsc --noEmit
npx tsx scripts/dateExpr.test.mts # date grammar round-trips against the REAL python resolver
python3 scripts/vendor_sb.py --check   # anti-drift gate
npm run build                     # re-vendors + next build

# sb/  (run from repo root)
python -m sb.demo                 # oracle round-trip
pytest sb/tests
python -m sb.conflicts --corpus corpus   # MONITOR-ONLY cross-storyline calendar time conflicts (never scored; see BACKLOG §2a)
python -m sb.probe                # synth a ~20-author stand-in corpus for sb.conflicts (real corpus is empty in git)
# oracle-score one authored node end-to-end:
#   schema.build_corpus([node]) -> scheduler.build_plan(...) -> engine.run(corpus, plan, oracle_model); ok iff score()==1.0

# benchmark run (requires Claude Code CLI installed + logged in; corpus/nodes/ must be populated)
./run.sh --model claude-sonnet-4-6 --seed 42 --days 300 > build/run.log 2>&1
python -m sb.analyze build/run.log --seed 42 --days 300
# to add filler noise (run the above first to get a clean baseline):
python -m sb.scale --filler 300 --seed 42 --days 300
./run.sh --model claude-sonnet-4-6 --corpus build/scaled --seed 42 --days 300 > build/run-scaled.log 2>&1
python -m sb.analyze build/run-scaled.log --corpus build/scaled --seed 42 --days 300
```

## Working agreements (standing decisions — don't relitigate without asking)

- **Don't change the grading engine (`sb/`)** for webapp work — make the authoring UI fit the engine, then re-vendor.
- Exact-day/time grading only; `within:Nd` is retired (field still exists in `sb` for back-compat, just unused).
- No free-slot / `no_overlap`; conflicts are authored reschedules. **Cross-storyline time double-booking is cosmetic under scoped grading (can't change a score) and is parked in BACKLOG §2a with the global-grading epic — `sb.conflicts` only measures it, never fixes it.** Don't add non-overlap to the serving algorithm without picking up that epic.
- Export-only; no import; no prod auth (all intentional).
- **Commit / PR style**: no em-dashes, no "Generated with Claude"/co-author attribution, conversational tone.
- Code formatting: dense, long lines fine, single-expression arrows on one line (see `~/.claude/CLAUDE.md`).
- Parked, do-not-act-yet: see `OPEN_QUESTIONS.md` (to-do grading details, deeper keyword ideas).

## Glossary

- **node / scenario / storyline** — one DAG of related emails (the unit you author).
- **obligation** — a named event/to-do; its name is also a node-scoped `@anchor`.
- **anchor** — a reusable named date (`{!name=…}` or an obligation name).
- **needle** — an email whose answer reuses an earlier anchor (the long-horizon retrieval test).
- **oracle** — the reference perfect-secretary solver; a valid corpus must oracle-solve 1.0.
- **tier** — difficulty guidance T1/T2/T3 in AUTHORING_GUIDE (≈ 30/40/30). Authoring guidance only; **not surfaced in the webapp UI**.
- **email type** — action vs no-action. The webapp derives this from the answer key (the "needs no action" checkbox = empty `ops`); the finer no_action/junk split (and the ≈ 80/12/7 mix) is authoring guidance, not tagged in the UI.

## Canonical docs (read these for depth)

- `GRADING_MODEL.md` — the grading contract (precise).
- `HOW_IT_WORKS.md` — the plain-language version + the worked "Project Atlas" example.
- `ANSWER_KEY_GRAMMAR.md` — the token/answer language.
- `AUTHORING_GUIDE.md` — how to author scenarios (type mix, tiers, needles).
- `docs/AUTHORING_WALKTHROUGH.md` — concrete click-by-click "author one storyline" example (the verified "VP onboarding" thread); covers focus mode + CEO-sent emails. The "don't break it" doc for club authors.
- `OPEN_QUESTIONS.md` — open design calls still being decided.
- `BACKLOG.md` — deferred features we'd build later (ranged offsets, global grading) + why + trigger to revisit.
- `docs/PROJECT_MAP.md`, `docs/RUNNING.md` — deeper engine/run/serving notes.
- `docs/MCP.md` — **stale** (describes an old FastAPI architecture; ignore).
