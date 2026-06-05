# Authoring UI: idiomatic redesign of date entry + consistent language

> Status: **DECIDED — Direction A. Phases 1–3 shipped; Blockly is gone** (branch
> `webapp-ui-idiomatic`). Only phase 4 (the optional natural-language accelerator) remains.
> Companion to `ANSWER_KEY_GRAMMAR.md` (the grammar this UI must produce) and the webapp
> README. Goal: make the authoring UI as idiomatic and low-effort as possible, with one
> consistent vocabulary for *writing emails* and *grading them*, by inheriting proven
> patterns instead of maintaining a bespoke Scratch/Blockly block editor.
>
> **Shipped:** `webapp/lib/dateExpr.ts` (structured parser/serializer, verified round-trip
> against `sb/resolver.py` via `webapp/scripts/dateExpr.test.mts`) and
> `webapp/components/DateBuilder.tsx` (the Direction-A fill-in-the-blank builder), wired into
> both the **answer key** (`AnswerKeyBuilder`) and the **email body** (`BodyEditor`, via an
> inline insert panel that adds the `{!name=}` emission).
>
> **Update (2026-06-05):** the author-facing raw "type it" escape hatch was REMOVED — every date
> is now composed through the structured builder (all eight bases are in the dropdown, incl.
> `week_of`/`month`). A raw text field survives only to display/repair a stored value the builder
> can't represent (legacy `in:`/`not_in`), never as a typing entry point. This closes the
> "two rules" friction §1 flags below: there is no longer any place that invites typing a date.
> The whole §3 vocabulary pass is in. The `blockly` dependency and `components/blockly/` are
> **deleted** — there is now one date UI everywhere. An incomplete pick (e.g. an anchor with no
> name chosen) is stored as empty rather than as invalid grammar, so it reads as "add the date."

## 1. The concern

The date-token builder is currently **Blockly** (the Scratch engine, Zelos renderer): a
full-screen modal where you drag and snap a custom block per grammar production
(`tk_serve`, `tk_anchor`, `tk_offset`, `tk_next`, `tk_this`, `tk_nth`, `tk_dom`,
`tk_week_of`, `tk_month`, `tk_monthref`), and a hand-written `blockToExpr` compiler turns
the block tree into the grammar string, which the real Python resolver then evaluates.

For a DSL whose expressions are *usually one to three tokens* (`+5d`, `next:THU`,
`@signing+2w`), a drag-and-drop canvas is heavy machinery. It also imposes a vocabulary
that doesn't match the grader's, and teaches one rule ("never type a date") that another
part of the same UI contradicts (the answer-key field accepts typed expressions).

## 2. What the grammar actually demands (why this is *not* a normal date picker)

A calendar widget or a pure natural-language parser is tempting but **cannot express our
grammar**, because our dates are *relative, compositional, and anchored*:

| Production | Example | Why a calendar / NL parser fails |
|---|---|---|
| `serve` base | `+5d` = "5 days after this email arrives" | the serve date isn't known at authoring time |
| `@anchor` base | `@signing+2w` | "the signing" is resolved at *serve* time from another email; you can't click it on a calendar |
| business days | `+3bd` | chrono-node and friends have no first-class business-day unit |
| `from <expr>` | `next:MON from @signing` | anchored weekday selection — not in any NL date lib |
| emission | `{!signing = +5d}` | publishes a *named* date for a later email — no off-the-shelf widget does this |
| intervals | `week_of:(serve+1w)`, `month:+1m` | a span, not a day |
| set predicates | `any_of:[…]`, `in … not_in @blackout` | answer-key only, no UI today |

**Conclusion:** the right control is a *structured builder that speaks our grammar*, with
a live resolver preview. Natural-language free text can be an **accelerator** for the
common serve-relative cases, but it cannot be the only path. (This is the gap an
NL-first redesign would hit — flagged here so we don't walk into it.)

## 3. Consistent language audit (writing ⇄ grading)

The grader/grammar names a concept one way; the UI often names it two or three other ways.
Pick one canonical author-facing word per concept and use it everywhere; show the
technical alias once, in the glossary only.

| Concept (grader term) | Said in the UI as… | Problem | Canonical author word |
|---|---|---|---|
| `node` | "storyline", "folder", "related email **thread**", "node id" | four words for one thing; Sidebar says "thread", everyone else "storyline" | **storyline** (alias: node) |
| obligation / op `name` | "thing", "action", "what to call this event", "name" | the load-bearing concept has no stable name | **name** (the obligation's name) |
| emit / `@anchor` | "publish", "anchor", "@name", "needle" | "needle" is a *different* idea (a retrieval test) mixed in with the *mechanism* | mechanism = **published date / @anchor**; keep **needle** only as the teaching word for the pattern |
| predicate `eq`/`by`/`in`/`any_of`/`not_in` | "on exactly (eq)", "on or before (by)", … | exposes the raw key to authors | plain English only: **on exactly / on or before / within / any of / not within** |
| token / `expr` | "date token", "block", "expression" | jargon, and tied to Blockly ("block") | **date** (a built date), shown as a chip |
| edge `static` | "static (fact, no deadline)" *and* "static (retrieval span)" | **same edge explained two different ways** (DependencyPicker vs DAG legend) | one line: **static — comes after, no deadline** |
| edge `date` | "date (carries a deadline)" / "date (deadline)" | OK, just unify | **date — comes after, carries a deadline** |
| tier | "difficulty", "tier", "T1/T2/T3" | fine | **difficulty (T1–T3)** |

**Correctness-risk teaching bug:** the body editor and guide say *"Never type a date by
hand — always the block builder,"* but `AnswerKeyBuilder` has a free-text predicate field
that the resolver validates live. The UI both forbids and allows typing. The redesign
resolves this: typing is fine **because the resolver validates it as you type** — the
guarantee was never "you can't type," it was "the date in the body and the answer key must
resolve to the same value." Say *that* instead.

## 4. Why Blockly is overkill here (grounded)

- **Blocks trade density and speed for syntax-safety.** Blockly's own design guide notes
  blocks are *lower density* (more screen per expression) and *higher viscosity* (small
  edits are harder) than text. That trade pays off for novices avoiding syntax errors in
  large programs — not for a 1–3 token date.
- **Coverage gaps today.** The toolbox has no block for the `{!name=}` emission (it's a
  separate checkbox), none for `any_of` / `not_in`, and the `from <expr>` clause is a
  buried optional input. So the "everything is a block" promise is already only partly
  kept, while the `blockToExpr` compiler must be hand-maintained against the grammar.
- **It interrupts writing.** A full-screen modal to drop in one date breaks the flow of
  composing an email; the research below favors *inline, in-context* entry.
- **Weight.** `blockly@^11` is hundreds of KB of editor for what is, semantically, a few
  dropdowns and a number field.

The literature points to **frame-based editing** (Kölling/Brown, *Stride*) as the proven
middle ground: "structured frames with slots for text" keep block-style safety while
matching text-entry speed — students made syntactic edits faster and spent less time in a
broken state than with either blocks or raw text.

## 5. Research-backed principles we should design to

- **Shrink both of Norman's gulfs.** *Execution:* the control should read like the English
  the author would write ("two weeks after the signing"). *Evaluation:* always show the
  resolved date ("→ Monday, August 17, 2026"). The Blockly modal *widens* the evaluation
  gulf — you must mentally compile blocks → grammar string → date. (Hutchins, Hollan &
  Norman, *Direct Manipulation Interfaces*; NN/g, *The Two UX Gulfs*.)
- **Recognition over recall** (Nielsen heuristic): offer the choices, don't make authors
  remember `nth:3,FRI,+1m`.
- **Hick's Law + progressive disclosure:** show only the relevant small set of choices at
  each step (base first; offsets and advanced predicates revealed on demand), instead of a
  flyout of 10 block types at once.
- **Live preview as you type** (Todoist, Dub): immediate, in-context feedback is the
  single highest-leverage affordance, and we already have the real resolver behind
  `/api/resolve` to power it.

## 6. Candidate directions

All three keep the **live Python-resolver preview** (the anti-drift guarantee) and a
**raw-expression escape hatch** validated by the resolver (so power users and the rare
full-grammar case — `any_of`, `not_in`, nested `from` — are never blocked).

### A. Fill-in-the-blank sentence builder *(recommended)*

A "Mad Libs" / natural-language form (Wroblewski): a sentence with small recognition-based
controls in the blanks. Inline, no modal. Inherits the *Natural Language Form* pattern and
*frame-based editing* research.

```
Date:  [ the day this email arrives ▾ ]  ( + [ 2 ] [ weeks ▾ ]  ✕ )  [ + add offset ]
                                                              → Monday, Aug 17, 2026
       base ▾ options:
         · the day this email arrives (serve)
         · a date another email set …  → [ @signing ▾ ]
         · next / this  [ Thursday ▾ ]   ( from [ @signing ▾ ]? )
         · the [ 3rd ▾ ] [ Friday ▾ ] of [ next month ▾ ]
         · day [ 25 ] of [ this month ▾ ]
         · the week of ( … )            ‹interval›
         · the whole month of [ … ]     ‹interval›
   ☐ other emails can refer to this date as  [ signing ]   ‹publishes @signing›
```

- **Maps to every production:** base dropdown = `serve | @anchor | next/this | nth | dom |
  week_of | month`; each "add offset" row = one `±N unit`; the checkbox = the `{!name=}`
  emission as a first-class control (fixing a current gap).
- **Pros:** reads like prose (small execution gulf); only the relevant choices visible
  (Hick); no drag, no modal, mobile/touch-friendly; drops the `blockly` dep and the
  `blockToExpr` compiler; one component reused for body tokens *and* answer-key predicates.
- **Cons:** building the inline control is real work (≈ the Blockly code it replaces, minus
  the dependency); deeply nested `from`/`week_of` expressions read awkwardly (→ escape hatch).
- **Effort:** **M**

### B. Smart text field with parse-on-blur + preview + picker fallback

Dub's "smart datetime picker" pattern, on chrono-node: type "in 2 weeks" / "next Thursday",
it parses on blur, shows a confirmed chip, falls back to a picker.

```
Date:  [ next thursday____________ ]  → Thu, Aug 13, 2026  ✓     [▦ pick]
       (couldn't parse "3 biz days after signing"? → opens builder A)
```

- **Pros:** fastest for the common serve-relative case; familiar (Todoist); tiny UI.
- **Cons (decisive):** chrono-node **cannot** express `@anchor`, `+Nbd`, `from @x`,
  `{!name=}`, `week_of`, `not_in` — i.e. exactly the parts that make *this* benchmark hard.
  As the only path it would silently cap authors at easy cases. **Only viable as an
  accelerator that pre-fills builder A**, not as the foundation.
- **Effort:** **S** (as an accelerator layer on top of A)

### C. Popover "date recipe" (base picker + offset rows + calendar quick-pick)

A compact popover combining a shadcn/ui-style base picker, offset rows, and Today/Tomorrow/
relative quick chips, built from inheritable components (Radix Popover/Select + react-day-picker).

```
┌ Build a date ───────────────────────────┐
│ Start from:  ( serve )( @anchor )( next…)│  ← segmented
│ Offsets:     + [2] [weeks ▾]   [+ row]   │
│ Quick:      [today][tomorrow][next Mon]  │
│ Preview:    → Monday, Aug 17, 2026       │
│ Publish as: ☐ [ signing ]                │
└──────────────────────────────────────────┘
```

- **Pros:** inherits maintained components instead of bespoke code; structured like A.
- **Cons:** a popover is still a small mode (less inline than A); the calendar grid is
  mostly dead weight since our dates are relative, not absolute.
- **Effort:** **M**

## 7. Recommendation

**Build A (the fill-in-the-blank sentence builder) as the primary control, with B layered
on as an optional accelerator and the validated raw-expression field kept as the escape
hatch.** Rationale: A is the only direction that expresses the *whole* grammar while
reading like the English an author already has in their head — it shrinks both gulfs, obeys
Hick/recognition/progressive-disclosure, and lets us delete the `blockly` dependency and
the hand-written compiler. B alone would quietly limit authors to the easy cases; as an
accelerator that pre-fills A, it's pure upside.

**Phasing**

1. **[DONE] Language pass (S, low-risk):** unified the vocabulary from §3 across Sidebar,
   EmailEditor primer, BodyEditor, AnswerKeyBuilder, DAG legend, ValidateBar, and the
   guide; fixed the two-explanations `static` edge and the "never type a date" contradiction.
2. **[DONE] Builder A for the answer-key predicate** (the higher-value, more error-prone
   slot), reusing the existing `/api/resolve` preview, with a validated raw-text escape
   hatch and an `any of` list. Covers eq/by/in/any_of; not_in + the in/not_in combo go
   through the escape hatch.
3. **[DONE] Builder A for body tokens** — an inline insert panel (no modal) reusing
   `DateBuilder` plus the `{!name=}` emission toggle and cursor insert; removed `TokenBlockly`
   + `blocks.ts` and dropped the `blockly` dependency.
4. **[TODO] Accelerator B** (chrono-node free-text → pre-fill A) as a fast path.

Each step is independently shippable and reversible; the resolver/grammar/grader are
untouched (anti-drift guarantee preserved end to end).

## Sources

- Hutchins, Hollan & Norman, *Direct Manipulation Interfaces* — https://www.lri.fr/~mbl/ENS/FONDIHM/2013/papers/Hutchins-HCI-85.pdf
- NN/g, *The Two UX Gulfs: Evaluation and Execution* — https://www.nngroup.com/articles/two-ux-gulfs-evaluation-execution/
- Interaction Design Foundation, *Gulf of Evaluation and Gulf of Execution* — https://www.interaction-design.org/literature/book/the-glossary-of-human-computer-interaction/gulf-of-evaluation-and-gulf-of-execution
- Hick's Law / progressive disclosure / recognition-vs-recall (NN/g + overviews) — https://www.geeksforgeeks.org/what-is-recognition-vs-recall-in-ux-design/ , https://www.thesigma.co/journal/hicks-law-ux
- Blockly Docs, *Block- vs. text-based languages* — https://docs.blockly.com/guides/design/languages/
- Kölling, Brown et al., *Frame-Based Editing: Combining the Best of Blocks and Text* — https://dl.acm.org/doi/10.1145/2818314.2818331 ; *Evaluation of a Frame-based Programming Editor* — https://dl.acm.org/doi/10.1145/2960310.2960319
- Weintrop & Wilensky, *Comparing Block-Based and Text-Based Programming* — https://www.cs.unm.edu/~learningcomputing/readings/17_weintrop_wilensky.pdf
- Luke Wroblewski / Codrops "Natural Language Form" ("Mad Libs" form); ui-patterns *Fill in the Blanks* — https://www.jroehm.com/2014/01/26/ui-pattern-natural-language-form/ , https://ui-patterns.com/patterns/FillInTheBlanks
- Todoist, *Introduction to dates and time* (natural-language input + live preview) — https://www.todoist.com/help/articles/introduction-to-dates-and-time-q7VobO
- chrono-node (MIT NL date parser) — https://github.com/wanasit/chrono
- Dub, *Building a Smart Datetime Picker — without using AI* (parse-on-blur + preview + picker fallback) — https://dub.co/blog/smart-datetime-picker
- shadcn/ui Date Picker (Popover + Calendar + react-day-picker; relative-date shortcuts) — https://ui.shadcn.com/docs/components/radix/date-picker
