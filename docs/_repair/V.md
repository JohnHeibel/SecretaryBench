# V — construct validity

Does the benchmark measure what it claims to measure? Scope: retrieval span, `sb/scale.py`,
tier reporting, and the binary per-email metric.

All measurements below were taken on 2026-08-17 against `corpus/` at HEAD (`67b3005`) with
`.venv/bin/python` (3.13). Every command is offline and free.

Two commands are cited repeatedly; both are reproduced here once.

**[CMD-FLOOR]** — score of a model that never calls a tool:
```bash
.venv/bin/python -c "
from datetime import date
from sb.schema import load_corpus
from sb.scheduler import build_plan, Levers
from sb.engine import Store, run
c=load_corpus('corpus'); p=build_plan(c,start_date=date(2026,6,1),seed=42,n_days=60,levers=Levers(1,5,7))
r=run(c,p,lambda e,b,x,s: None, store=Store(c))
print('null floor', r.passed,'/',r.total,'=',f'{r.score():.1%}')
n=sum(len(e.body)+len(e.subject)+len(e.sender) for e in c.emails.values())
print('corpus text', n, 'chars ~', n//4, 'tokens')"
# -> null floor 64 / 167 = 38.3%
# -> corpus text 43474 chars ~ 10868 tokens
```

**[CMD-TRACE]** — retained tool calls per recorded run:
```bash
.venv/bin/python -c "
import re,collections
D=re.compile(r'── day (\d+) · .+? · (\d+) new email')
for f in ['outputs/opus.md','outputs/sonnet.md','past/claude-haiku-4-5.md']:
    cur=None; tot=collections.Counter(); bad=0; days=0; em=0
    for l in open(f,encoding='utf-8',errors='replace'):
        m=D.search(l)
        if m: cur=int(m.group(2)); days+=1; em+=cur; continue
        if cur is not None and l.strip().startswith('tools'):
            t=l.split('tools',1)[1].split(chr(128269))[0]
            g=t.count('get_email'); tot.update(x.strip() for x in t.split(','))
            if g<cur: bad+=1
    print(f,'days',days,'emails',em,'get_email',tot['get_email'],'search_inbox',tot['search_inbox'],'deficit days',bad)"
# -> outputs/opus.md           days 57 emails 167 get_email 166 search_inbox 3 deficit days 1
# -> outputs/sonnet.md         days 57 emails 167 get_email  57 search_inbox 3 deficit days 46
# -> past/claude-haiku-4-5.md  days 16 emails 167 get_email  40 search_inbox 0 deficit days 15
```

---

## V-1 the corpus is two orders of magnitude too small to push any fact out of context
Status: open
Severity: blocks-measurement
Cost to verify: free-offline

**What's wrong.** `sb/span.py:4-6` states the research claim: a large span means "the fact has
likely scrolled out of context by the time it's needed, so the model must `search_inbox` to
recover it." The entire authored corpus is 43,474 characters of email text (~10.9k tokens), and
the largest gap between a needle's setup email and its payoff email is 18,389 characters of
body (~4.6k tokens). Every model in the roster has a context window of at least 200k tokens, so
no authored fact has ever scrolled out of context in any recorded run, and `search_inbox` has
never been necessary to answer a single email.

**Evidence.**
- The claim: `sb/span.py:4-6`; `TIER_LIST.md:14-15` ("**Retrieval span** — how long ago the
  needed fact was set"); `TIER_LIST.md:110-116` (T3 Hardener A: "far enough it has scrolled out
  of the model's window, so it MUST `search_inbox`").
- Corpus text size and the null floor: **[CMD-FLOOR]** → `corpus text 43474 chars ~ 10868 tokens`.
- Needle gap, measured as the email text served strictly between emitter and payoff in the
  seed-42 / `daily_max=5` plan (the plan the recorded runs used):

```bash
.venv/bin/python -c "
import itertools
from datetime import date
from sb.schema import load_corpus
from sb.scheduler import build_plan, Levers
from sb.span import spans
c=load_corpus('corpus'); p=build_plan(c,start_date=date(2026,6,1),seed=42,n_days=60,levers=Levers(1,5,7))
order=[e for b in p.per_day for e in b]; pos={e:i for i,e in enumerate(order)}
cum=list(itertools.accumulate(len(c.emails[e].body) for e in order))
g=[cum[pos[r['needs']]]-cum[pos[r['from']]] for r in spans(c,p)]
print('needle gap (body chars): n',len(g),'mean',sum(g)//len(g),'max',max(g))
b=[len(e.body) for e in c.emails.values()]; print('mean body chars', sum(b)//len(b))"
# -> needle gap (body chars): n 24 mean 7065 max 18389
# -> mean body chars 227
```
  Counting subject, sender and a 120-char JSON envelope per served record raises the mean gap to
  11,899 chars (~3.0k tokens), the maximum to 31,063 chars (~7.8k tokens), and the whole-corpus
  figure to 63,514 chars (~15.9k tokens).
- The mean authored email body is 227 characters. Fifty-seven days of this do not fill a window.
- The recorded runs behaved accordingly: **[CMD-TRACE]** shows `search_inbox` retained 3 times in
  opus, 3 in sonnet, 0 in haiku — all of opus's and sonnet's on day 1
  (`outputs/opus.md:8`, `outputs/sonnet.md:8` are the only lines in either file containing the
  string, per `grep -n search_inbox outputs/*.md past/*.md`).

**Why it matters for benchmark validity.** The stated research contribution is retrieval span.
On the corpus as authored and as run, retrieval span is not a variable — every fact is verbatim
in the transcript when it is needed. Whatever the 54% / 54% / 59% scores measure, it is not
retrieval under context pressure, and no honest paper claim about long-span retrieval can be
supported by any run recorded so far.

**Options.**
1. Scale the corpus with filler until the gap exceeds a real window, and re-run. Cost: this is
   the only path that tests the actual claim; but see V-5 — at the filler levels that work, the
   do-nothing floor rises to 60-68%, and it requires a paid full run per data point.
2. Re-scope the claim to what the corpus supports: multi-day *dependency* reasoning within
   context (mean day-span 10.8, max 28 — `python -m sb.span`), not out-of-context retrieval.
   Cost: free and immediately honest; abandons the differentiating claim.
3. Force retrieval mechanically rather than by volume — e.g. have the harness elide or compact
   older turns so the setup email is provably absent when the payoff arrives. Cost: makes span a
   controlled independent variable at any corpus size; changes the harness contract and would
   need its own validation.
4. Keep the corpus and report span as a covariate with an explicit null result ("accuracy did not
   vary with span; span never exceeded ~8k tokens"). Cost: cheap and defensible; the headline
   becomes a negative result.

**Overlaps with:** V-2, V-5, V-7, K-*.

**Open questions.** The exact session token count is unknowable from the artifacts — the
transcript is discarded at `sb/live/runner.py:513-516` and there is no `--out` (O-*). The
argument above is therefore a bound on the *gap*, not a measurement of the *window occupancy*.
Does the Claude CLI's auto-compaction ever trigger on a 57-turn session of this size? If it does,
retrieval could be forced by compaction rather than by span, which would be a different (and
uncontrolled) independent variable.

---

## V-2 the "search_inbox on 1 of 57 days" figure is real for opus and unmeasurable for sonnet and haiku
Status: open
Severity: blocks-measurement
Cost to verify: free-offline

**What's wrong.** §2.8 of the brief reports search rates that §2.7 simultaneously undermines,
because `sb/live/runner.py:183` collapses assistant messages by id and `sb/analyze.py:51-52`
derives `searched` purely from the surviving `tools` string. Counting retained `get_email` calls
against emails served settles it per run: opus's trace is complete (166 retained vs 167 emails;
1 deficit day of 57), sonnet's and haiku's are not (57 vs 167 and 40 vs 167; 46 and 15 deficit
days). So opus genuinely searched on exactly one day, sonnet's and haiku's search behaviour is
unrecoverable from the record, and no run's search rate can be attributed per email in any case.

**Evidence.**
- The loss mechanism: `sb/live/runner.py:183` `seen[msg.get("id", id(ev))] = msg`, then
  `:187-194` iterates `seen.values()`. A message id is therefore represented by whichever
  assistant event arrived last, so tool calls are lost exactly when a model emits several
  `tool_use` blocks under one message id.
- The lower bound that settles it: the day prompt carries no bodies — `runner.py:418-424` POSTs
  bodies to the store and `runner.py:426-429` sends only the date. The only ways to read a body
  are `get_email` (`sb/live/store_app.py:225-230`) and `search_inbox`
  (`sb/live/store_app.py:192-213`). So a model that acts on N emails in a day must call
  `get_email` at least N times.
- **[CMD-TRACE]**: opus retained `get_email` 166 for 167 emails, with only **1 of 57** days
  showing fewer retained `get_email` than emails served — and that day is day 1, the one day it
  searched (`outputs/opus.md:8`: `ToolSearch, list_new_emails, get_email, get_email,
  search_inbox, search_inbox, get_email, get_email, search_inbox, create_event, create_event,
  create_event` for 5 emails). On the other 56 days retained `get_email` equals emails served
  exactly, so those days' traces already account for a full serial pass over the mail and no
  `search_inbox` can be hiding inside them.
- Sonnet retained exactly one `get_email` per day on all 57 days (57 total) regardless of whether
  1 or 5 emails arrived, and only 41 create-type calls (`create_event` 22 + `create_todo` 19)
  against 48 `matched` + 5 `cancelled` outcomes (V-4's tally). A flat 1-per-day cannot be a
  complete transcript; sonnet batched parallel tool calls and the dict kept the last block of
  each batch. Full retained-tool histograms:
```bash
.venv/bin/python -c "
import re,collections
D=re.compile(r'── day (\d+) · .+? · (\d+) new email')
for f in ['outputs/opus.md','outputs/sonnet.md','past/claude-haiku-4-5.md']:
    cur=None; tot=collections.Counter()
    for l in open(f,encoding='utf-8',errors='replace'):
        m=D.search(l)
        if m: cur=int(m.group(2)); continue
        if cur is not None and l.strip().startswith('tools'):
            tot.update(x.strip() for x in l.split('tools',1)[1].split(chr(128269))[0].split(','))
    print(f, dict(tot.most_common()))"
# opus:   get_email 166, list_new_emails 57, create_event 54, create_todo 35, update_todo 13,
#         delete_event 6, update_event 6, search_inbox 3, ToolSearch 1   (341 total)
# sonnet: list_new_emails 57, get_email 57, create_event 22, create_todo 19, update_event 6,
#         search_inbox 3, delete_event 3, update_todo 3, ToolSearch 1    (171 total)
# haiku:  get_email 40, list_new_emails 16, create_todo 11, create_event 7, ToolSearch 4 (78 total)
```
- Haiku: 40 retained `get_email` for 167 emails over 16 days (`daily_max=21`), 15 of 16 deficit
  days, and 18 retained create-type calls against 49 `matched` + 5 `cancelled`. Same verdict as
  sonnet.
- None of the three runs lost a day to errors
  (`grep -c "ERROR after retries" outputs/opus.md outputs/sonnet.md past/claude-haiku-4-5.md`
  → 0 for each) and no day printed `tools  (none)`, so the deficits are not turn failures.
- The unit is wrong even when the trace is right: `sb/live/runner.py:57-63` prints one `tools`
  line per **day**, and `sb/analyze.py:36-41` says so in its own docstring — `searched` means
  "the model used search_inbox at least once during this email's day."
- The signal is additionally pinned to zero by construction. A needle's payoff must be served
  after its emitter, so no needle can fall on day 1; day 1 is the only day either model searched.
  Running the retrieval analysis on either recorded run gives `searched% 0%` in **every** span
  bin:
```bash
.venv/bin/python -m sb.analyze outputs/opus.md --corpus corpus --seed 42 --days 60 --daily-max 5
# span bin   n  accuracy searched%
# 0-50      19      37%        0%
# 50-100     5      60%        0%
```

**Why it matters for benchmark validity.** The single number that would show the benchmark
exercises retrieval — "how often did the model search when it needed an old fact" — is 0% for
every needle in both full runs, for a structural reason, and is untrustworthy for two of the
three models for an unrelated instrumentation reason. Fixing the trace (O-*) will make the number
measurable but cannot make it non-zero: that requires V-1.

**Options.**
1. Fix `_parse_stream` to append rather than overwrite, and re-run. Cost: makes future search
   rates trustworthy for all models; does not recover sonnet's or haiku's, and the existing three
   logs stay uninterpretable on this axis.
2. Instrument retrieval on the server side instead of the client side — `store_app.py` already
   sees every `/inbox/search` hit and already has a warning channel (`store_app.py:192-200`), so
   a per-call log there is driver-independent and cannot be lost to a parser. Cost: needs a
   persisted artifact (O-*); adds a store endpoint.
3. Attribute retrieval per email rather than per day by making the turn boundary one email
   instead of one day (`DAY_LOOP_DESIGN_ISSUE.md` exists on this). Cost: exact attribution;
   changes the benchmark's core loop and multiplies turn count and spend.
4. Drop `searched` from the reported analysis entirely and report only accuracy-by-span. Cost:
   free, removes a column that is currently guaranteed to read 0%; loses the mechanism story.

**Overlaps with:** O-*, V-1, V-8.

**Open questions.** Does the Claude CLI emit one `assistant` event per content block (the
mechanism assumed above) or one per completed message? The opus/sonnet split is fully explained by
the former plus a difference in serial-vs-parallel tool use, but this was inferred from the logs,
not confirmed against a captured stream — no raw stream was saved. Capturing one turn of raw
`stream-json` would settle it definitively and is the cheapest possible live check.

---

## V-3 38.3% of the score is available to a model that never calls a tool
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** A null model that takes no action on any email scores 64/167 = 38.3%, because
all 56 no-action emails pass (`sb/grader.py:186-195`) and all 8 cancel-only emails also pass —
`sb/grader.py:155-158` passes a `cancel` when nothing matching is on the calendar, which is
trivially true if the object was never created. The reported scores are quoted against an implied
floor of 0%, so 54% reads as "half right" when it is 15.6 points above a floor requiring no
capability at all. On the 103 emails that actually require action, opus scores 33.0%, sonnet
35.0% and haiku 36.9%.

**Evidence.**
- **[CMD-FLOOR]** → `null floor 64 / 167 = 38.3%`. Decomposition (56 no-action + 8 cancel-only,
  no others):
```bash
.venv/bin/python -c "
from datetime import date
from sb.schema import load_corpus
from sb.scheduler import build_plan, Levers
from sb.engine import Store, run
c=load_corpus('corpus'); p=build_plan(c,start_date=date(2026,6,1),seed=42,n_days=60,levers=Levers(1,5,7))
res=run(c,p,lambda e,b,x,s: None, store=Store(c)).results
ok={k for k,v in res.items() if v.passed}
noop={e.id for e in c.emails.values() if not e.answer.ops}
canc={e.id for e in c.emails.values() if e.answer.ops and all(o.verb=='cancel' for o in e.answer.ops)}
print(len(ok),'=',len(ok&noop),'no-action +',len(ok&canc),'cancel-only + other',len(ok-noop-canc))"
# -> 64 = 56 no-action + 8 cancel-only + other 0
```
- Split of each recorded run against that floor (join the log's PASS/FAIL rows with the null
  model's per-email verdicts):

| run | raw | on the 64 a null model passes | on the 103 requiring action | normalised above floor |
|---|---|---|---|---|
| `outputs/opus.md` | 90/167 = 53.9% | 56/64 = 87.5% | **34/103 = 33.0%** | 25.2% |
| `outputs/sonnet.md` | 91/167 = 54.5% | 55/64 = 85.9% | **36/103 = 35.0%** | 26.2% |
| `past/claude-haiku-4-5.md` | 98/167 = 58.7% | 60/64 = 93.8% | **38/103 = 36.9%** | 33.0% |

  (normalised = `(score − 0.383) / (1 − 0.383)`.) Reproduce with:
```bash
.venv/bin/python -c "
import re
from datetime import date
from sb.schema import load_corpus
from sb.scheduler import build_plan, Levers
from sb.engine import Store, run
c=load_corpus('corpus'); p=build_plan(c,start_date=date(2026,6,1),seed=42,n_days=60,levers=Levers(1,5,7))
res=run(c,p,lambda e,b,x,s: None, store=Store(c)).results
ok={k for k,v in res.items() if v.passed}
ROW=re.compile(r'\b(PASS|FAIL)\b\s+\[(\d+)\]\s+(\S+)')   # same regex as sb/analyze.py:25
for f in ['outputs/opus.md','outputs/sonnet.md','past/claude-haiku-4-5.md']:
    r={m.group(3):m.group(1)=='PASS' for m in
       (ROW.search(l) for l in open(f,encoding='utf-8',errors='replace')) if m}
    easy=[e for e in r if e in ok]; hard=[e for e in r if e not in ok]
    s=sum(r.values())/len(r); fl=len(ok)/len(res)
    print(f'{f:<28} raw {sum(r.values())}/{len(r)}={s:.1%}'
          f'  easy {sum(r[e] for e in easy)}/{len(easy)}'
          f'  hard {sum(r[e] for e in hard)}/{len(hard)}={sum(r[e] for e in hard)/len(hard):.1%}'
          f'  norm {(s-fl)/(1-fl):.1%}')"
# -> outputs/opus.md            raw 90/167=53.9%  easy 56/64  hard 34/103=33.0%  norm 25.2%
# -> outputs/sonnet.md          raw 91/167=54.5%  easy 55/64  hard 36/103=35.0%  norm 26.2%
# -> past/claude-haiku-4-5.md   raw 98/167=58.7%  easy 60/64  hard 38/103=36.9%  norm 33.0%
```
- `sb/grader.py:155-158`: `passed = len(title_set) == 0` for `cancel`, with `title_set` drawn from
  the cumulative node pool (`:151-152`). No check that the object ever existed.
- 24 of the 167 emails (14.4%) are needles at all — see V-1's span command. None of the 24 is in
  the null model's pass set, so the needle subset is the one part of the score with a genuine 0%
  floor.

**Why it matters for benchmark validity.** The headline number is not comparable to a naive
reading, and the 5-point spread between the three models understates the actual spread: 25.2% /
26.2% / 33.0% normalised is a 7.8-point gap, still with haiku on top. Any future "we improved the
score from 54% to X" is uninterpretable without stating the floor, because a grader change that
touches no-action or cancel handling moves 64 of 167 points.

**Options.**
1. Report the null-model floor alongside every score, and report accuracy on the actionable
   subset separately. Cost: free, one line of harness output; does not change the benchmark, only
   how it is read.
2. Make `cancel` falsifiable — require that a matching object existed in the node before the
   cancel op's serve date. Cost: removes 8 free points and makes the verb mean something; it is a
   grader change, so it must not ship in the same commit as a corpus change (`CLAUDE.md`).
3. Rebalance the corpus so no-action emails are a stated fraction rather than an emergent 33.5%.
   Cost: gives the floor a designed value; a corpus change, and the "over-action is the most
   common mistake" framing in `sb/live/runner.py:101` argues for keeping no-action well
   represented.
4. Score by op rather than by email, so the 56 no-action emails contribute 56 of 190 judgements
   rather than 56 of 167 points. Cost: see V-4; changes the denominator and breaks comparability
   with all recorded runs.

**Overlaps with:** G-*, V-4, K-*.

**Open questions.** Is 38.3% the right floor to publish, or should the floor be an "acts on
everything obvious, ignores everything subtle" baseline that would score higher? Nobody has built
a second baseline. §4.2 of the brief (the true score is unknown) is unaffected by this finding —
the floor bounds the *interpretation*, not the *truth*, of the score.

---

## V-4 190 op-level judgements are compressed into 167 binary points
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** `sb/grader.py:197-198` sets `passed = all(d["passed"] for d in details)` and
`sb/live/runner.py:519` sums those booleans, so a 4-op email is worth the same one point as a
"read an FYI and do nothing" email, and a 3-of-4 answer scores identically to a 0-of-4 answer. The
corpus contains 134 graded ops plus 56 no-action checks = 190 judgements, collapsed to 167 points.
The §2.3 failure taxonomy survives in the printed log but not in the metric, and
`sb/grader.py:202` keeps only the *first* failing reason in `headline`.

**Evidence.**
- `sb/grader.py:24` ("Output: EmailResult(passed, max=1, details[]). Binary per email."),
  `:197-198`, `:202`; `sb/live/runner.py:519`.
- Answer shapes and the 190 count:
```bash
.venv/bin/python -c "
import collections
from sb.schema import load_corpus
c=load_corpus('corpus')
h=collections.Counter(len(e.answer.ops) for e in c.emails.values())
print('ops-per-email',dict(sorted(h.items())),'total ops',sum(k*v for k,v in h.items()))"
# -> ops-per-email {0: 56, 1: 94, 2: 12, 3: 4, 4: 1} total ops 134
```
  56 no-action details + 134 op details = 190, which matches the log exactly: each of the three
  logs contains 190 `why` lines (`grep -c '     why' outputs/opus.md` → 190).
- The taxonomy **is** recoverable from the printed log, and reproduces §2.3 exactly. Normalising
  the reason strings (`no <kind> titled like` → one bucket, `found N matching` → one bucket,
  `over-acted…` → one bucket, `should be cancelled…` → one bucket) gives, for
  `outputs/opus.md`: 52 / 51 / 41 / 18 / 14 / 6 / 5 / 3 — identical to the brief's opus column;
  sonnet 47 / 51 / 48 / 18 / 12 / 5 / 5 / 4 and haiku 52 / 56 / 49 / 7 / 17 / 5 / 0 / 4 likewise
  match. So the brief's §2.3 reproduces; what is lost is *causal* resolution inside the 52-case
  bucket (§4.3), not the tally.
- It is not recoverable from any machine artifact, because there is none: the `details` list is
  printed at `sb/live/runner.py:80-85` and then discarded with the store at `:513-516`.

**Why it matters for benchmark validity.** A single binary per email makes the score a weighted
mixture of three incommensurable tasks — no-action recognition (56 points), single-op execution
(94 points) and all-or-nothing multi-op execution (17 points) — with weights nobody chose. It also
means the metric cannot separate "the model did the hard part and fumbled one op" from "the model
did nothing", which is precisely the distinction the failure analysis needs.

**Options.**
1. Keep the binary email score as the headline and additionally emit per-op counts and the reason
   bucket. Cost: no comparability break; needs the structured output from O-*.
2. Switch the denominator to ops (190 judgements). Cost: gives every judgement equal weight and
   makes multi-op emails count for what they contain; all four recorded runs become
   non-comparable, and a per-op score raises the null floor's arithmetic again (V-3).
3. Report both, with the email score as primary. Cost: honest and cheap; two numbers to explain,
   and readers will quote whichever is higher.
4. Split the reported score by answer shape (no-action / single-op / multi-op) rather than
   aggregating. Cost: exposes the mixture directly; three numbers with n=56/94/17, the last too
   small to be stable.

**Overlaps with:** G-*, O-*, V-3.

**Open questions.** Should a partially-correct multi-op email score partial credit, or is
all-or-nothing the right semantics for an obligation set? The register has no stated position, and
the answer determines whether option 2 is a fix or a different benchmark.

---

## V-5 `sb/scale.py` is unusable as run and self-defeating where it works
Status: open
Severity: blocks-measurement
Cost to verify: free-offline

**What's wrong.** `sb/scale.py` exists specifically to force retrieval distance
(`sb/scale.py:1-19`), and no recorded run used it. It fails outright at several filler levels
(30, 60 and 200 all raise `InfeasibleSchedule`) while 90, 120 and 150 succeed, so the difficulty
knob is not monotonic. Where it does work, every filler email is a graded no-action email
(`sb/scale.py:80`: `"answer": {"ops": []}`), so the do-nothing floor of V-3 rises from 38.3% to
59.9% / 64.1% / 67.5% — above every score any real model has ever recorded.

**Evidence.**
- Feasibility sweep:
```bash
for f in 30 60 90 120 150 200; do echo "### filler=$f"; \
  .venv/bin/python -m sb.scale --filler $f --seed 42 --days 300 --dst build/sc_$f 2>&1 | tail -3; done
```
  `filler=30` → `can't serve 30 filler in 300 days: 2 email(s) never served ...
  ['Innovation-comp.one-pager-looks-good', 'Innovation-comp.quick-favor-before-the-final']`;
  `filler=60` and `filler=200` → same failure naming the five `Company-Retreat.*` emails;
  `filler=90/120/150` → oracle 100% with needle span mean 39.8 / 56.6 / 62.6.
  The error's own advice is wrong: `sb/scale.py:113-114` says "raise `--days` (… try >= 96)"
  when `--days 300` was already supplied and the real cause is over-constrained serve windows.
  Raising the window does not help: at `--days 400`, `--filler 175` succeeds (needle span mean
  **59.5**, *lower* than 150's 62.6 — non-monotonic in span as well as in feasibility) while
  `--filler 200` fails with the identical five-email message.
- Floor at each working level (null model on the scaled corpus, `--days 300`):

| corpus | emails | no-action | null floor | needle span (emails) | max needle gap |
|---|---|---|---|---|---|
| `corpus/` | 167 | 56 (34%) | **38.3%** | mean 31.6 / max 83 | ~7.8k tok |
| `build/sc_90` | 257 | 146 (57%) | **59.9%** | mean 39.8 / max 190 | ~102k tok |
| `build/sc_120` | 287 | 176 (61%) | **64.1%** | mean 56.6 / max 182 | ~123k tok |
| `build/sc_150` | 317 | 206 (65%) | **67.5%** | mean 62.6 / max 186 | ~135k tok |

  (same null-model recipe as **[CMD-FLOOR]** with `load_corpus('build/sc_N')` and `n_days=300`;
  gaps computed as in V-1 with subject/sender and a 120-char envelope.)
- Filler emails are 4,627 body characters on average (~1,157 tokens) versus 227 for authored
  emails — `sb/scale.py:76-78` deliberately generates 16-22 paragraphs. So ~173 filler emails
  between a setup and its payoff would be needed to clear a 200k window, which is inside the
  band where feasibility becomes erratic (175 works, 200 does not).
- The runner grades every email in the plan including filler (`sb/live/runner.py:503-511` iterates
  `batch`, with no filter on id), so filler is not free scenery — at `--filler 120` it is 120
  extra guaranteed-passable points on a 287-email denominator.

**Why it matters for benchmark validity.** The one instrument that could establish the headline
claim of V-1 cannot currently be pointed at it: the settings that produce real out-of-context
distance also make two-thirds of the score obtainable by doing nothing, and the settings between
them fail to schedule. A run at `--filler 120` would produce a number that looks like a benchmark
score and means "the model correctly ignored 176 newsletters."

**Options.**
1. Exclude filler from grading (tag the `gen-filler` node and skip it in the runner's per-email
   loop). Cost: restores the floor at any filler level; the model can then be scored only on
   authored emails while still paying the context cost, but filler stops testing over-action.
2. Weight or subsample filler in the score instead of excluding it. Cost: keeps some over-action
   pressure; introduces a weighting parameter that has to be justified.
3. Fix the scheduler's serve-window handling so filler scales monotonically, before deciding
   anything about grading. Cost: makes the knob usable at all; `Innovation-comp` and
   `Company-Retreat` windows are the blockers and touching them is a corpus change (K-*).
4. Abandon volume-based scaling in favour of an explicit context-eviction mechanism in the harness
   (V-1 option 3). Cost: decouples span from corpus size entirely; larger harness change.

**Overlaps with:** V-1, V-3, K-*.

**Open questions.** Why does `--filler 30` fail while `--filler 90` succeeds? The scheduler's
interaction with `depends_on` windows is not documented, and until it is, no filler level can be
argued to be "the" setting. Also unmeasured: whether a ~124k-token scaled corpus would trigger CLI
auto-compaction, which would change what "out of context" means mid-run.

---

## V-6 `tier` is loaded and never read; the by-tier report lives behind a file nothing writes
Status: open
Severity: blocks-measurement
Cost to verify: free-offline

**What's wrong.** `TIER_LIST.md:175-181` asks for a score-by-tier report and states "The score
that matters is T3", and every email in the corpus is tagged (50 T1 / 67 T2 / 50 T3). `Email.tier`
is declared at `sb/schema.py:81` and parsed at `sb/schema.py:230`, and then read by nothing.
`sb/analyze.py:94-105` looks for tiers in `<corpus>/needles.json` under a key `reasoning_tier`,
a file that does not exist anywhere in the repo and that no code writes — so the tier column of
the only tier-aware tool always prints `untagged`.

**Evidence.**
- `grep -rn "\.tier\b" --include="*.py" . | grep -v "^./.venv"` → **zero hits**. The field is
  declared (`sb/schema.py:81`, `tier: str | None = None`) and populated (`sb/schema.py:230`,
  `tier=raw.get("tier")`), and no attribute access to it exists anywhere in the tree. It is
  write-only.
- `find . -name needles.json -not -path "./.venv/*"` → no output.
  `grep -rn "needles.json\|reasoning_tier\|tier_name" . --include="*.py" --include="*.json"
  --include="*.md" | grep -v "^./.venv"` → four hits, all inside `sb/analyze.py` (`:94`, `:104`,
  `:105`, `:118`). Nothing produces the file; `sb/scale.py` does not write it either
  (`sb/scale.py:85-96` copies nodes and writes only `gen_filler.json`).
- Observed output of the tier report on a real run:
```bash
.venv/bin/python -m sb.analyze outputs/opus.md --corpus corpus --seed 42 --days 60 --daily-max 5
# tier          0-50       50-100       100+
# untagged    37% (19)    60% (5)        ·
```
- The tags do exist:
```bash
.venv/bin/python -c "
import collections
from sb.schema import load_corpus
print(collections.Counter(e.tier for e in load_corpus('corpus').emails.values()))"
# -> Counter({'T2': 67, 'T1': 50, 'T3': 50})
```
- The `100+` span bin (`sb/analyze.py:29`) is structurally unreachable on the authored corpus:
  max span is 83.

**Why it matters for benchmark validity.** `TIER_LIST.md:180` names T3 accuracy as the number
that ranks models, and that number has never been computed for any run. The benchmark's own
difficulty design is therefore untested against its own results — the three models could differ
sharply on T3 and identically on T1/T2 and nobody would know.

**Options.**
1. Read `corpus.emails[eid].tier` directly in `sb/analyze.py` and delete the `needles.json`
   branch. Cost: makes the tier report exist for every email, not just the 24 needles; small edit,
   no artifact needed.
2. Emit a `needles.json` manifest from `sb/scale.py` so the existing branch works as designed.
   Cost: keeps the tool's shape; adds a generated file to keep in sync with the corpus, and still
   only covers needles.
3. Report tier accuracy from the runner itself rather than from the log-scraper, once structured
   output exists (O-*). Cost: the tier breakdown becomes a first-class run artifact; blocked on
   phase 1.
4. Drop the tier concept from the tooling and rely on span alone. Cost: removes dead code and a
   dead doc promise; discards the only difficulty axis that is actually tagged (see V-7).

**Overlaps with:** O-*, C-*, V-7.

**Open questions.** Is the `needles.json` branch a vestige of a generator that was removed, or a
manifest that was always intended to be hand-written? If the former, there may be other analysis
paths that silently degrade to `untagged`.

---

## V-7 the authored tier gradient does not exist on the axis it is defined by
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** `TIER_LIST.md:23-27` and `:191` define the T1→T2→T3 gradient primarily by
retrieval distance: "same email" → "a few days back" → "weeks back (out of context)". Measured on
the corpus, T3's needles are *closer* than T2's on both axes (email-span 31.1 vs 33.3; day-span
10.4 vs 11.8), and 32 of the 50 T3 emails have no cross-email anchor reference in their answer
key at all. Thirteen of the 50 T3 emails are pure no-action, which `TIER_LIST.md:57-59` assigns to
T1 unless the bait is genuinely tricky.

**Evidence.**
```bash
.venv/bin/python -c "
import collections
from datetime import date
from sb.schema import load_corpus
from sb.scheduler import build_plan, Levers
from sb.span import spans
c=load_corpus('corpus'); p=build_plan(c,start_date=date(2026,6,1),seed=42,n_days=60,levers=Levers(1,5,7))
recs=spans(c,p); by=collections.defaultdict(list)
for r in recs: by[c.emails[r['needs']].tier].append((r['email_span'],r['day_span']))
for t,v in sorted(by.items()):
    e=[a for a,_ in v]; d=[b for _,b in v]
    print(t,'n',len(v),'email_span mean %.1f max %d'%(sum(e)/len(e),max(e)),
          '| day_span mean %.1f max %d'%(sum(d)/len(d),max(d)))
ids={r['needs'] for r in recs}
for t in ['T1','T2','T3']:
    g=[e for e in c.emails if c.emails[e].tier==t]
    print(t,len(g),'emails;',len([e for e in g if e in ids]),'with a foreign-anchor answer;',
          len([e for e in g if not c.emails[e].answer.ops]),'pure no-action')"
# -> T2 n 6  email_span mean 33.3 max 65 | day_span mean 11.8 max 22
# -> T3 n 18 email_span mean 31.1 max 83 | day_span mean 10.4 max 28
# -> T1 50 emails; 0 with a foreign-anchor answer; 19 pure no-action
# -> T2 67 emails; 6 with a foreign-anchor answer; 24 pure no-action
# -> T3 50 emails; 18 with a foreign-anchor answer; 13 pure no-action
```
- `TIER_LIST.md:110-116` requires T3 Hardener A to place the fact "far enough it has scrolled out
  of the model's window"; the T3 maximum is 83 emails ≈ 7.8k tokens (V-1).
- `TIER_LIST.md:166-168` requires every T3 to stack "≥2" hardeners; only 18 of 50 T3 emails carry
  even one measurable one (a foreign anchor reference). Hardeners B and C are not machine-visible,
  so this is a lower bound, not a refutation of the other two.

**Why it matters for benchmark validity.** T1/T2/T3 is the benchmark's stated difficulty
construct and the intended basis for model ranking. On the one hardener that can be measured
today, the tiers are indistinguishable and slightly inverted. Any by-tier report built after V-6
would therefore be reporting a gradient that is not in the corpus, which is worse than reporting
none.

**Options.**
1. Re-derive tiers from measured span rather than author intent, and report both. Cost: makes the
   tier axis honest and automatic; discards the authored judgement about task ambiguity, which is
   the part span cannot see.
2. Keep author tiers and add a separate measured-span axis, reporting the crosstab. Cost:
   preserves both constructs and exposes the mismatch instead of hiding it; two axes to explain
   with n=24 on the span side.
3. Re-author the T3 set to satisfy Hardener A, using `sb/scale.py` to place the distance. Cost:
   the design's intent is met; a corpus change (K-*) gated on V-5 being usable.
4. Retire Hardener A from the tier definition and rest T3 on Hardeners B and C (multi-fact
   constraint and under-action bait), which the corpus may already satisfy. Cost: the tier
   definition becomes checkable against what exists; abandons the retrieval framing at the tier
   level, which then has to be dropped from V-1's claim too.

**Overlaps with:** K-*, V-1, V-6.

**Open questions.** How many T3 emails satisfy Hardeners B or C? Neither is machine-detectable
from the schema, so this needs a manual pass over the 50 T3 emails — the same pass phase 1.5
already schedules for the hand-grade, and worth combining.

---

## V-8 the span axis of the analysis is reconstructed from levers the artifact never recorded
Status: open
Severity: distorts-measurement
Cost to verify: free-offline

**What's wrong.** `sb/analyze.py:88-90` rebuilds the serve plan from `--seed`, `--days` and the
three lever flags, then computes span from that reconstruction (`:91`), because the run itself
saved nothing but printed text. The recorded logs carry no config stamp — `outputs/opus.md` ends
at `SCORE 90/167 (54%)` with no lever line, since M-2's stamp postdates them. Guessing the levers
wrong silently rebins every needle rather than erroring.

**Evidence.**
- `tail -4 outputs/opus.md` shows the score line and nothing else; the levers are inferred only
  from the 57-day count matching `daily_max=5` (register, "Corpus health check").
- Same log, same seed, one lever changed:
```bash
.venv/bin/python -m sb.analyze outputs/opus.md --corpus corpus --seed 42 --days 60 --daily-max 5
#   0-50: n 19, 37%   50-100: n 5, 60%   100+: —
.venv/bin/python -m sb.analyze outputs/opus.md --corpus corpus --seed 42 --days 60 --daily-max 21
#   0-50: n 13, 23%   50-100: n 10, 60%   100+: n 1, 100%
```
  Overall needle accuracy is 42% either way (the per-email verdicts come from the log), but the
  entire x-axis of the retrieval finding changes, including whether the top bin is populated at
  all.
- The tool's own documented invocation (`sb/analyze.py:11`) defaults to `--corpus build/scaled
  --days 300` with no lever flags, i.e. `daily_max=5` — which would silently mismatch any run made
  at `daily_max=21`, such as `past/claude-haiku-4-5.md`.
- Both recorded runs show accuracy *rising* with span (opus 37%→60%, sonnet 42%→60%), the opposite
  of the hypothesis in `sb/analyze.py:2-4`. Recorded here as tool output only, not as a result:
  n=19 and n=5, and the bin membership is a function of the guessed lever (see above).

**Why it matters for benchmark validity.** The independent variable in the benchmark's headline
plot is not measured; it is recomputed after the fact from parameters supplied by hand at analysis
time. Nothing in the pipeline detects a mismatch, so a wrong flag produces a plausible-looking
graph rather than an error.

**Options.**
1. Have the runner emit the span of every email alongside its verdict, so the analysis reads
   recorded spans instead of recomputing them. Cost: the x-axis becomes an observation; requires
   the structured output of O-*.
2. Stamp seed, levers and a corpus hash into the run artifact (M-2 already stamps them into the
   printed footer) and have `sb/analyze.py` parse and enforce them, refusing to run on a mismatch.
   Cost: cheap, catches the error class; leaves the pre-phase-0 logs unanalysable.
3. Pin one canonical lever set for all published runs and drop the flags from `sb/analyze.py`.
   Cost: removes the failure mode entirely; loses `daily_max` as an experimental variable, and
   `past/claude-haiku-4-5.md` used a different one.
4. Leave as is and document the required flags per artifact in the register. Cost: free; keeps a
   silent-wrong-answer path open for the next person.

**Overlaps with:** C-*, O-*, V-2.

**Open questions.** Were the recorded opus and sonnet runs definitely at `daily_max=5`? The
57-day match is strong but circumstantial, and everything in V-2's and V-8's span numbers depends
on it. A corpus hash in the artifact would also be needed — `past/claude-sonnet-4-5.md:4` records
`sha 809d389794dd79a9` for a 176-email corpus that no longer exists, so its spans can never be
recomputed at all.

---

## Reproduction notes on the evidence brief

**Reproduced exactly.**
- §2.1 scores: opus 90/167, sonnet 91/167, haiku 98/167 (scraped from the logs' PASS/FAIL rows
  with `sb/analyze.py:25`'s own regex).
- §2.3 failure tally: all three columns match after normalising the free-text reason strings
  (V-4 evidence). Each log contains exactly 190 `why` lines = 56 no-action + 134 op judgements.
- §2.8 search counts: opus 1 of 57 days, sonnet 1 of 57, haiku 0 of 16 (**[CMD-TRACE]**;
  `grep -n search_inbox outputs/*.md past/*.md` hits exactly one line in `outputs/opus.md` and
  one in `outputs/sonnet.md`, both day 1, and none in `past/claude-haiku-4-5.md`).
- §2.8 span: mean 31.6, max 83, n=24 (`.venv/bin/python -m sb.scale --filler 0 --seed 42
  --days 200 --dst build/scaled0` → `needle span: max 83, mean 31.6 (n=24)`,
  `oracle: 167/167 = 100%`).
- §2.7 tool-call totals: sonnet 57 `get_email`, opus 166 (**[CMD-TRACE]**).
- §2.9 "`analyze.py` never reads `email.tier`": confirmed, and sharpened in V-6 — it reads a
  different key from a file nothing writes.
- §2.10 corpus state: 167 emails, 134 graded ops, 50 T1 / 67 T2 / 50 T3.

**Could not reproduce as stated.**
- **§2.7's sonnet day-1 quote is attributed to the wrong artifact.** The brief quotes
  "`list_new_emails, get_email, search_inbox, search_inbox, create_event` for 8 emails". That line
  is `past/claude-sonnet-4-5.md:18` — the retired 176-email corpus, 19 days, `daily_max=21`
  (`past/claude-sonnet-4-5.md:4`). `outputs/sonnet.md:8` reads `ToolSearch, list_new_emails,
  get_email, search_inbox, search_inbox, create_todo, search_inbox` over **5** emails. The
  lossiness conclusion still holds; the citation does not.
- **§2.7's "Same harness, so those counts are not comparable" is the wrong inference.** The counts
  are directly comparable and are the most useful diagnostic available: they differ because opus
  serialised its tool calls and sonnet batched them, which is exactly what distinguishes a
  faithful trace from a collapsed one (V-2).

**Listed in §4 as not established; now established (or bounded).**
- **The §2.7↔§2.8 confound is settled.** `search_inbox on 1 of 57 days` is a genuine behavioural
  fact for opus and an artifact-limited non-measurement for sonnet and haiku. Method: retained
  `get_email` equals emails served on 56 of opus's 57 days versus a flat 1-per-day for sonnet on
  all 57, against a hard lower bound of one `get_email` per email read (`runner.py:426-429`,
  `store_app.py:225-230`). Command: **[CMD-TRACE]**.
- **A hard floor now bounds §4.2.** The true score remains unknown, but 64 of 167 points
  (38.3%) require no capability, and the three runs score 33.0% / 35.0% / 36.9% on the 103 emails
  that require action. Any hand-grade in phase 1.5 should sample the actionable 103, not the 167.
- **§4.5 is sharpened, not refuted.** Normalised above the floor, the three runs are 25.2% /
  26.2% / 33.0% — the convergence is real and haiku still leads, but the gap is 7.8 points rather
  than the 5 the raw scores suggest.
- **New, not in the brief:** the corpus cannot exercise retrieval at any setting currently usable
  (V-1, V-5), and the `searched` column of the retrieval analysis is pinned to 0% for every needle
  by construction (V-2).
